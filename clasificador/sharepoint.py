"""Volcado del extracto a una carpeta de SharePoint, via Microsoft Graph.

Es el hermano de ``sheets.py``: mismo extracto, mismo formato, otro destino.

Diferencia de fondo con Google Sheets: alla la planilla se edita in situ (se
leen las filas y se agregan las que faltan). Aca el destino es un **archivo**,
no una base editable por API, asi que el ciclo es bajar -> mergear -> subir.
Hay dos razones por las que no se usa la Excel REST API de Graph, que si
permitiria editar celdas sueltas:

* declara "Application: Not supported" en range-update y createSession, o sea
  que no funciona con token de aplicacion (sin eso no hay corrida desatendida);
* no expone validacion de datos en ninguna version, asi que no podria crear ni
  reparar el desplegable de Categoria.

Autenticacion (ver README, seccion SharePoint):

1. Client credentials (app-only): ``SP_TENANT_ID`` + ``SP_CLIENT_ID`` +
   ``SP_CLIENT_SECRET``. Es el unico modo realmente desatendido.
2. Device code: ``SP_TENANT_ID`` + ``SP_CLIENT_ID``, sin secreto. Pide loguearse
   en el navegador la primera vez y deja el token en una cache local. Sirve para
   probar o para correrlo a mano; el refresh token caduca, no lo pongas en cron.
"""

from __future__ import annotations

import base64
import io
import os
import time
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import quote, unquote, urlsplit

import openpyxl
import requests

from clasificador.excel import (
    HOJA_DATOS,
    Movimiento,
    claves_con_ocurrencia,
    leer_movimientos,
)
from clasificador.libro import armar_libro, combinar, leer_libro_existente
from clasificador.reglas import JuegoDeReglas

GRAPH = "https://graph.microsoft.com/v1.0"
AUTORIDAD = "https://login.microsoftonline.com"

# En app-only el scope es SIEMPRE el literal ".default": los permisos efectivos
# salen de lo que un admin consintio en el registro de la app, no del pedido.
ALCANCE_APP = ["https://graph.microsoft.com/.default"]
# En device code, en cambio, los scopes son concretos y sin el prefijo del recurso.
ALCANCE_DELEGADO = ["Files.ReadWrite.All"]

TIPO_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
NOMBRE_POR_DEFECTO = "Extractos_cobranzas.xlsx"
CACHE_DEVICE_CODE = Path.home() / ".cache" / "cobranzas" / "token_sharepoint.json"

# El PUT simple admite hasta 250 MB, pero arriba de unos pocos MB conviene la
# upload session: parte el archivo y cada trozo se reintenta solo, sin resubir
# los que ya entraron. El extracto pesa kilobytes, asi que en la practica
# siempre va por el PUT simple.
LIMITE_SUBIDA_SIMPLE = 4 * 1024 * 1024
# La doc exige que cada trozo sea multiplo de 320 KiB y menor a 60 MiB.
TAMANIO_TROZO = 10 * 320 * 1024

# Timeout partido: si el host no contesta en 15s el problema es la red y no
# tiene sentido esperar; la lectura, en cambio, puede tardar lo que tarde una
# subida. Con un solo numero el script se quedaba 5 minutos mudo.
TIMEOUT = (15, 300)
# Reintentos de los errores transitorios (throttling y caidas momentaneas de
# Graph). Son pocos a proposito: si el tenant esta limitando de verdad,
# insistir no ayuda.
REINTENTOS = 3
ESTADOS_REINTENTABLES = (429, 503, 504)
ESPERA_MAXIMA = 60

# Significado de los ':x:' del link compartido. No esta documentado por
# Microsoft, sale de observar los links; solo se usa para avisar temprano si el
# usuario pego un link de archivo en vez de uno de carpeta.
TIPOS_LINK = {
    "f": "carpeta",
    "x": "excel",
    "w": "word",
    "p": "powerpoint",
    "b": "pdf",
    "i": "imagen",
    "v": "video",
    "u": "generico",
}
# En las URL de SharePoint '/s/' es la forma corta de '/sites/'.
PREFIJOS_SITIO = {"s": "sites", "sites": "sites", "t": "teams", "teams": "teams"}


# --- link compartido -------------------------------------------------------
def codificar_link(url: str) -> str:
    """Codifica una URL compartida como share id de Graph: ``u!`` + base64url.

    Algoritmo de la doc de ``/shares``: base64 de la URL completa (con su
    ``?e=...``), sacar el relleno ``=``, ``/`` -> ``_``, ``+`` -> ``-``.
    """
    b64 = base64.b64encode(url.encode("utf-8")).decode("ascii")
    return "u!" + b64.rstrip("=").replace("/", "_").replace("+", "-")


def decodificar_link(share_id: str) -> str:
    """Inversa de ``codificar_link``. Existe para poder testear el ida y vuelta."""
    if not share_id.startswith("u!"):
        raise ValueError("Un share id codificado empieza con 'u!'.")
    cuerpo = share_id[2:].replace("_", "/").replace("-", "+")
    cuerpo += "=" * (-len(cuerpo) % 4)
    return base64.b64decode(cuerpo).decode("utf-8")


@dataclass(frozen=True)
class LinkCompartido:
    """Las partes utiles de un link del tipo ``https://host/:f:/s/Sitio/token``."""

    hostname: str
    tipo: str
    ruta_sitio: str
    url: str


def analizar_link(url: str) -> LinkCompartido:
    """Descompone un link de SharePoint sin pegarle a la red.

    Sirve para dos cosas: avisar si el link no es de carpeta, y tener la ruta
    del sitio por si hay que ir por el camino ``/sites`` (ver ``resolver_sitio``).
    """
    partes = urlsplit(url)
    if not partes.hostname:
        raise ValueError(f"No parece una URL de SharePoint: {url!r}")
    segmentos = [unquote(s) for s in partes.path.split("/") if s]

    tipo = "desconocido"
    if segmentos and segmentos[0].startswith(":") and segmentos[0].endswith(":"):
        tipo = TIPOS_LINK.get(segmentos[0].strip(":")[:1], "desconocido")
        segmentos = segmentos[1:]
    if segmentos and segmentos[0] in {"r", "e"}:  # /r/ solo lectura, /e/ edicion
        segmentos = segmentos[1:]

    ruta_sitio = ""
    if len(segmentos) >= 2 and segmentos[0] in PREFIJOS_SITIO:
        ruta_sitio = f"/{PREFIJOS_SITIO[segmentos[0]]}/{segmentos[1]}"

    return LinkCompartido(partes.hostname, tipo, ruta_sitio, url)


# --- autenticacion ---------------------------------------------------------
@dataclass(frozen=True)
class ConfigSharePoint:
    """Credenciales de Entra ID. Sin ``client_secret`` se usa device code."""

    tenant_id: str
    client_id: str
    client_secret: str | None = None

    @classmethod
    def desde_entorno(cls) -> "ConfigSharePoint":
        faltan = [v for v in ("SP_TENANT_ID", "SP_CLIENT_ID") if not os.environ.get(v)]
        if faltan:
            raise RuntimeError(
                "Faltan variables de entorno para SharePoint: "
                + ", ".join(faltan)
                + ".\n"
                "  SP_TENANT_ID     ID del directorio (Entra ID > Overview).\n"
                "  SP_CLIENT_ID     ID de la aplicacion registrada.\n"
                "  SP_CLIENT_SECRET secreto de la app; sin el se usa device code\n"
                "                   (login en el navegador, no sirve para cron).\n"
                "Ver la seccion 'Volcar a SharePoint' del README."
            )
        return cls(
            tenant_id=os.environ["SP_TENANT_ID"],
            client_id=os.environ["SP_CLIENT_ID"],
            client_secret=os.environ.get("SP_CLIENT_SECRET") or None,
        )


def _token_app_only(config: ConfigSharePoint) -> str:
    """Client credentials: sin usuario, sin ventana de navegador, para cron."""
    import msal

    app = msal.ConfidentialClientApplication(
        config.client_id,
        authority=f"{AUTORIDAD}/{config.tenant_id}",
        client_credential=config.client_secret,
    )
    # MSAL cachea el token en memoria y lo renueva solo cuando hace falta.
    resp = app.acquire_token_for_client(scopes=ALCANCE_APP)
    if "access_token" not in resp:
        raise RuntimeError(_explicar_error_token(resp))
    return resp["access_token"]


def _token_device_code(config: ConfigSharePoint, cache: Path = CACHE_DEVICE_CODE) -> str:
    """Flujo delegado: imprime un codigo y espera a que alguien se loguee.

    Guarda el refresh token en ``cache`` para no pedir login en cada corrida.
    Igual caduca (ventana deslizante de ~90 dias) y se rompe si cambia la
    contrasena o una politica de acceso condicional: no es desatendido.
    """
    import msal

    almacen = msal.SerializableTokenCache()
    if cache.exists():
        _restringir_permisos(cache)  # por si quedo flojo de una corrida vieja
        almacen.deserialize(cache.read_text(encoding="utf-8"))

    app = msal.PublicClientApplication(
        config.client_id,
        authority=f"{AUTORIDAD}/{config.tenant_id}",
        token_cache=almacen,
    )

    resp = None
    cuentas = app.get_accounts()
    if cuentas:
        resp = app.acquire_token_silent(ALCANCE_DELEGADO, account=cuentas[0])

    if not resp:
        flujo = app.initiate_device_flow(scopes=ALCANCE_DELEGADO)
        if "user_code" not in flujo:
            raise RuntimeError(
                "No pude iniciar el device code flow: "
                f"{flujo.get('error_description') or flujo}.\n"
                "Revisa que la app tenga 'Allow public client flows' = Yes en la "
                "pestana Authentication del registro de Entra ID."
            )
        print(flujo["message"], flush=True)
        resp = app.acquire_token_by_device_flow(flujo)  # bloquea hasta el login

    if almacen.has_state_changed:
        _guardar_cache(cache, almacen.serialize())

    if "access_token" not in resp:
        raise RuntimeError(_explicar_error_token(resp))
    return resp["access_token"]


def _restringir_permisos(cache: Path) -> None:
    """Deja el cache y su directorio legibles solo por el dueno.

    Se llama tambien al leer: un archivo que quedo con 0644 de una corrida
    vieja no se arregla solo, y adentro hay un refresh token delegado que con
    Files.ReadWrite.All vale por todo lo que ve el usuario durante ~90 dias.
    """
    cache.parent.chmod(0o700)
    cache.chmod(0o600)


def _guardar_cache(cache: Path, contenido: str) -> None:
    """Escribe el cache del token naciendo ya con 0600.

    Crear el archivo y recien despues hacerle chmod deja una ventana en la que
    el refresh token es legible por cualquier usuario de la maquina; en un host
    compartido (el servidor del cron, Cloud Shell) eso alcanza. ``mkdir`` con
    ``mode`` solo aplica cuando crea, asi que el directorio se corrige aparte.
    """
    cache.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    cache.parent.chmod(0o700)
    fd = os.open(cache, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w", encoding="utf-8") as archivo:
        archivo.write(contenido)
    cache.chmod(0o600)  # si el archivo ya existia, el modo de os.open no aplica


def _explicar_error_token(resp: dict) -> str:
    """Traduce los errores de MSAL que mas se ven al pedir el token."""
    codigo = resp.get("error", "")
    detalle = resp.get("error_description", "") or str(resp)
    ayuda = ""
    if codigo == "invalid_client":
        ayuda = (
            "\nEl secreto es invalido o esta vencido. Los client secrets caducan "
            "(normalmente a los 6 o 24 meses): genera uno nuevo en Entra ID > la "
            "app > Certificates & secrets."
        )
    elif codigo in {"unauthorized_client", "invalid_request"}:
        ayuda = (
            "\nRevisa que SP_CLIENT_ID y SP_TENANT_ID sean los de la pestana "
            "Overview del registro de la app."
        )
    elif "AADSTS65001" in detalle or codigo == "consent_required":
        ayuda = (
            "\nFalta el consentimiento de administrador. Un Global Administrator "
            "del tenant tiene que entrar a Entra ID > la app > API permissions y "
            "apretar 'Grant admin consent'."
        )
    return f"No pude obtener el token de Microsoft Graph ({codigo}): {detalle}{ayuda}"


def obtener_token(config: ConfigSharePoint | None = None) -> str:
    """Devuelve un access token de Graph, por el modo que corresponda."""
    config = config or ConfigSharePoint.desde_entorno()
    if config.client_secret:
        return _token_app_only(config)
    return _token_device_code(config)


# --- cliente ---------------------------------------------------------------
@dataclass(frozen=True)
class Carpeta:
    """Destino del volcado. Cachea estos dos ids: el link puede caducar."""

    drive_id: str
    item_id: str
    nombre: str = ""
    url_web: str = ""


@dataclass(frozen=True)
class ArchivoRemoto:
    id: str
    nombre: str
    tamanio: int
    modificado: str
    etag: str = ""


class ClienteGraph:
    """Lo minimo de Microsoft Graph que necesita este proyecto.

    Se le puede inyectar la sesion para testear sin red.
    """

    def __init__(self, token: str, sesion: requests.Session | None = None) -> None:
        self.sesion = sesion or requests.Session()
        self.sesion.headers["Authorization"] = f"Bearer {token}"

    @classmethod
    def desde_entorno(cls) -> "ClienteGraph":
        return cls(obtener_token())

    def pedir(
        self, metodo: str, ruta: str, ok_404: bool = False, **kwargs
    ) -> requests.Response | None:
        """Llama a Graph y convierte los errores en mensajes accionables.

        ``ok_404=True`` devuelve ``None`` en lugar de reventar: es lo que
        distingue "el archivo todavia no existe" de un error de verdad.

        Los 429/503/504 se reintentan solos: decirle "esperá un minuto" a una
        corrida de cron es decirle nada, el volcado del dia simplemente no pasa.
        """
        url = ruta if ruta.startswith("http") else GRAPH + ruta
        for intento in range(REINTENTOS):
            resp = self._request(metodo, url, **kwargs)
            if resp.status_code not in ESTADOS_REINTENTABLES:
                break
            if intento == REINTENTOS - 1:
                break
            espera = _espera_sugerida(resp, intento)
            print(
                f"Graph respondio {resp.status_code} y pidio esperar {espera}s; "
                f"reintentando ({intento + 2} de {REINTENTOS})...",
                flush=True,
            )
            time.sleep(espera)
        if resp.status_code == 404 and ok_404:
            return None
        if resp.status_code >= 400:
            raise RuntimeError(_explicar_error_graph(metodo, url, resp))
        return resp

    def _request(self, metodo: str, url: str, **kwargs) -> requests.Response:
        """El pedido crudo, con las fallas de red traducidas."""
        try:
            return self.sesion.request(metodo, url, timeout=TIMEOUT, **kwargs)
        except requests.exceptions.RequestException as error:
            raise RuntimeError(_explicar_error_red(error, metodo, url)) from error

    def json(self, metodo: str, ruta: str, **kwargs) -> dict:
        return self.pedir(metodo, ruta, **kwargs).json()


def _espera_sugerida(resp: requests.Response, intento: int) -> int:
    """Segundos a esperar: los que diga ``Retry-After``, o backoff con tope."""
    cabecera = (resp.headers or {}).get("Retry-After", "")
    try:
        return min(int(cabecera), ESPERA_MAXIMA)
    except (TypeError, ValueError):
        return min(2 ** (intento + 1), ESPERA_MAXIMA)


def _url_segura(url: str) -> str:
    """Tapa el share id antes de imprimir una URL.

    ``/shares/u!<base64url>`` lleva adentro el link compartido entero, con su
    ``?e=...``, y se decodifica en una linea (ver ``decodificar_link``). Un 404
    que termina en el log del cron o pegado en un ticket no tiene por que
    regalar el link vivo de la carpeta. El resto de las rutas (``/sites``,
    ``/drives``) no tienen secretos y se muestran enteras.
    """
    inicio = url.find("/shares/")
    if inicio < 0:
        return url
    fin = url.find("/", inicio + len("/shares/"))
    return url[:inicio] + "/shares/<link-redactado>" + (url[fin:] if fin >= 0 else "")


def _explicar_error_red(error: Exception, metodo: str, url: str) -> str:
    """Traduce las fallas de red, que en una PC de oficina son las mas comunes.

    El orden importa: en requests, ``ProxyError`` y ``SSLError`` heredan de
    ``ConnectionError``, y ``ConnectTimeout`` de las dos ramas.
    """
    fallas = requests.exceptions
    if isinstance(error, fallas.ProxyError):
        ayuda = (
            "No pude salir por el proxy. Si la empresa usa uno, exportalo antes "
            "de correr el comando:\n"
            "  export HTTPS_PROXY=http://usuario:clave@proxy.empresa:8080"
        )
    elif isinstance(error, fallas.SSLError):
        ayuda = (
            "Fallo la verificacion del certificado TLS. Pasa cuando el firewall "
            "inspecciona el trafico: pedile a sistemas el certificado de la "
            "empresa y exporta REQUESTS_CA_BUNDLE=/ruta/al/certificado.pem. "
            "No desactives la verificacion."
        )
    elif isinstance(error, fallas.Timeout):
        ayuda = "El servidor no contesto a tiempo. Reintenta en un rato."
    elif isinstance(error, fallas.ConnectionError):
        ayuda = (
            "No hay conexion con graph.microsoft.com. Revisa internet y el DNS; "
            "si estas detras de un proxy, exporta HTTPS_PROXY."
        )
    else:
        ayuda = "Revisa la conexion y reintenta."
    return (
        f"No pude hablar con Microsoft Graph ({metodo} {_url_segura(url)}).\n"
        f"{ayuda}\n{type(error).__name__}: {error}"
    )


def _explicar_error_graph(metodo: str, url: str, resp: requests.Response) -> str:
    """Mensajes de error que digan que permiso falta y quien lo concede."""
    try:
        error = resp.json().get("error", {})
        detalle = f"{error.get('code')}: {error.get('message')}"
    except ValueError:
        detalle = resp.text[:400]

    ayuda = ""
    if resp.status_code == 401:
        ayuda = (
            "\nEl token no sirve o no tiene permisos consentidos. Un Global "
            "Administrator tiene que apretar 'Grant admin consent' en Entra ID > "
            "la app > API permissions."
        )
    elif resp.status_code == 403:
        ayuda = (
            "\nLa app esta autenticada pero no tiene acceso a este sitio. Con el "
            "permiso Sites.Selected hacen falta DOS pasos y seguro falta el "
            "segundo: ademas del consentimiento, un SharePoint Administrator "
            "tiene que otorgarle el rol 'write' sobre el sitio, con\n"
            "  POST /v1.0/sites/{siteId}/permissions\n"
            '  {"roles":["write"],"grantedToIdentities":'
            '[{"application":{"id":"<client-id>","displayName":"<nombre>"}}]}\n'
            "o con PnP PowerShell: Grant-PnPAzureADAppSitePermission "
            "-AppId <client-id> -Permissions Write -Site <url del sitio>."
        )
    elif resp.status_code == 404:
        # El consejo depende del camino: por link no hay nombres que revisar, y
        # por ruta de sitio no hay ningun link que pueda haber caducado.
        if "/shares/" in url:
            ayuda = (
                "\nGraph no encuentra la carpeta del link. Revisa que el link "
                "siga vigente (el '?e=...' se puede revocar) y que la app tenga "
                "permiso sobre el sitio; con Sites.Selected sin el grant, Graph "
                "responde 404 en vez de 403."
            )
        else:
            ayuda = (
                "\nGraph no encuentra el recurso. Por este camino no hay ningun "
                "link de por medio: lo que suele estar mal es lo que se escribio "
                "a mano. Revisa el nombre exacto del sitio (--sitio), el de la "
                "biblioteca (--biblioteca) y el de la subcarpeta (--subcarpeta); "
                "ojo que la biblioteca principal se llama 'Documents' o "
                "'Documentos' segun el idioma del tenant. Con Sites.Selected sin "
                "el grant sobre el sitio, Graph tambien responde 404 en vez de 403."
            )
    elif resp.status_code == 412:
        ayuda = (
            "\nAlguien modifico el archivo en SharePoint mientras corria esto. "
            "Volve a correr el comando: baja la version nueva y la mergea."
        )
    elif resp.status_code == 423:
        ayuda = (
            "\nEl archivo esta bloqueado porque alguien lo tiene abierto en Excel "
            "de escritorio. Pedile que lo cierre y reintenta."
        )
    elif resp.status_code == 429:
        ayuda = (
            "\nGraph te esta limitando (throttling) y los reintentos automaticos "
            f"no alcanzaron ({REINTENTOS} intentos). Espera unos minutos."
        )
    elif resp.status_code in (500, 503, 504):
        ayuda = (
            "\nEs una falla momentanea del lado de Microsoft, no de la "
            f"configuracion. Se reintento {REINTENTOS} veces sin suerte: volve "
            "a correr el comando mas tarde."
        )
    return f"Graph {resp.status_code} en {metodo} {_url_segura(url)}\n{detalle}{ayuda}"


# --- resolucion de la carpeta ----------------------------------------------
def resolver_carpeta(cliente: ClienteGraph, url: str) -> Carpeta:
    """Resuelve un link compartido de carpeta (':f:') a drive_id + item_id.

    Conviene correrlo UNA vez y guardar los dos ids: el link lleva un token de
    comparticion (``?e=...``) que se puede revocar o caducar, los ids no.
    """
    link = analizar_link(url)
    if link.tipo not in ("carpeta", "desconocido"):
        raise RuntimeError(
            f"El link apunta a un archivo ({link.tipo}), no a una carpeta. "
            "Compartí la carpeta desde SharePoint y pegá ese link."
        )
    share_id = quote(codificar_link(url), safe="")
    item = cliente.json(
        "GET",
        f"/shares/{share_id}/driveItem",
        params={"$select": "id,name,webUrl,folder,parentReference"},
    )
    if "folder" not in item:
        raise RuntimeError(
            f"El link resuelve a {item.get('name')!r}, que no es una carpeta."
        )
    drive_id = (item.get("parentReference") or {}).get("driveId", "")
    if not drive_id:
        raise RuntimeError(
            "Graph devolvio la carpeta sin parentReference.driveId. Pasa cuando "
            "el link apunta a la raiz de una biblioteca: usa --sitio y --ruta en "
            "lugar de --carpeta-url."
        )
    return Carpeta(drive_id, item["id"], item.get("name", ""), item.get("webUrl", ""))


def resolver_sitio(
    cliente: ClienteGraph,
    hostname: str,
    ruta_sitio: str,
    subcarpeta: str = "",
    biblioteca: str | None = None,
) -> Carpeta:
    """Resuelve la carpeta por ruta de sitio, sin depender del link compartido.

    Es el camino estable para produccion, y el unico que funciona con el permiso
    Sites.Selected: ``/shares`` no lo acepta.
    """
    sitio = cliente.json(
        "GET", f"/sites/{hostname}:{'/' + ruta_sitio.strip('/')}?$select=id,webUrl"
    )
    if biblioteca:
        drives = cliente.json("GET", f"/sites/{sitio['id']}/drives?$select=id,name")
        elegido = next(
            (d for d in drives["value"] if d["name"].casefold() == biblioteca.casefold()),
            None,
        )
        if elegido is None:
            nombres = ", ".join(repr(d["name"]) for d in drives["value"])
            raise RuntimeError(
                f"No hay una biblioteca llamada {biblioteca!r} en el sitio. Hay: {nombres}."
            )
        drive_id = elegido["id"]
    else:
        # La biblioteca por defecto se llama 'Documents' o 'Documentos' segun el
        # idioma del tenant; /drive la devuelve sin tener que adivinar el nombre.
        drive_id = cliente.json("GET", f"/sites/{sitio['id']}/drive?$select=id")["id"]

    campos = "?$select=id,name,webUrl"
    if subcarpeta.strip("/"):
        ruta = quote(subcarpeta.strip("/"))
        # ok_404 para poder decir que carpetas SI hay, en lugar del 404 pelado.
        resp = cliente.pedir(
            "GET", f"/drives/{drive_id}/root:/{ruta}{campos}", ok_404=True
        )
        if resp is None:
            raise RuntimeError(_sin_subcarpeta(cliente, drive_id, subcarpeta))
        item = resp.json()
    else:
        item = cliente.json("GET", f"/drives/{drive_id}/root{campos}")
    return Carpeta(drive_id, item["id"], item.get("name", ""), item.get("webUrl", ""))


def _sin_subcarpeta(cliente: ClienteGraph, drive_id: str, subcarpeta: str) -> str:
    """El 404 de ``--subcarpeta``, con la lista de lo que si hay en la raiz."""
    hijos = cliente.json(
        "GET", f"/drives/{drive_id}/root/children", params={"$select": "name,folder"}
    )
    carpetas = [h["name"] for h in hijos.get("value", []) if "folder" in h]
    hay = ", ".join(repr(n) for n in carpetas) if carpetas else "ninguna"
    return (
        f"No hay una carpeta {subcarpeta!r} en la biblioteca. En la raiz hay: {hay}.\n"
        "Ojo que --subcarpeta es la ruta DENTRO de la biblioteca, sin el nombre "
        "de la biblioteca adelante: 'Extractos', no "
        "'Documentos compartidos/Extractos'."
    )


# --- archivos --------------------------------------------------------------
def listar(cliente: ClienteGraph, carpeta: Carpeta) -> list[ArchivoRemoto]:
    """Lista los archivos de la carpeta, siguiendo la paginacion de Graph."""
    ruta = f"/drives/{carpeta.drive_id}/items/{carpeta.item_id}/children"
    # 'folder' tiene que estar en el $select: sin pedirla, Graph no devuelve la
    # faceta y el filtro de abajo no filtra nada, asi que las subcarpetas se
    # colarian como si fueran archivos.
    params = {"$select": "id,name,size,lastModifiedDateTime,eTag,folder", "$top": "200"}
    archivos = []
    while ruta:
        pagina = cliente.json("GET", ruta, params=params)
        for item in pagina.get("value", []):
            if "folder" in item:
                continue
            archivos.append(
                ArchivoRemoto(
                    id=item["id"],
                    nombre=item["name"],
                    tamanio=int(item.get("size") or 0),
                    modificado=item.get("lastModifiedDateTime", ""),
                    etag=item.get("eTag", ""),
                )
            )
        # El nextLink ya trae los parametros adentro: reenviarlos los duplica.
        ruta, params = pagina.get("@odata.nextLink", ""), None
    return archivos


def buscar(cliente: ClienteGraph, carpeta: Carpeta, nombre: str) -> ArchivoRemoto | None:
    for archivo in listar(cliente, carpeta):
        if archivo.nombre.casefold() == nombre.casefold():
            return archivo
    return None


def descargar(cliente: ClienteGraph, carpeta: Carpeta, item_id: str) -> bytes:
    return cliente.pedir(
        "GET", f"/drives/{carpeta.drive_id}/items/{item_id}/content"
    ).content


def subir(
    cliente: ClienteGraph,
    carpeta: Carpeta,
    nombre: str,
    datos: bytes,
    etag: str | None = None,
) -> dict:
    """Crea o reemplaza el archivo en la carpeta.

    ``etag``: si se pasa, el PUT falla con 412 cuando alguien toco el archivo
    desde que lo bajamos, en lugar de pisarle el trabajo en silencio.
    """
    if len(datos) <= LIMITE_SUBIDA_SIMPLE:
        return _subir_simple(cliente, carpeta, nombre, datos, etag)
    return _subir_en_partes(cliente, carpeta, nombre, datos)


def _subir_simple(
    cliente: ClienteGraph,
    carpeta: Carpeta,
    nombre: str,
    datos: bytes,
    etag: str | None,
) -> dict:
    cabeceras = {"Content-Type": TIPO_XLSX}
    if etag:
        cabeceras["if-match"] = etag
    return cliente.json(
        "PUT",
        f"/drives/{carpeta.drive_id}/items/{carpeta.item_id}:/{quote(nombre)}:/content",
        # En el PUT simple conflictBehavior va en la URL; en la upload session,
        # adentro del body. 'replace' mantiene el item id, asi que los links que
        # la gente ya tiene guardados siguen funcionando.
        params={"@microsoft.graph.conflictBehavior": "replace"},
        headers=cabeceras,
        data=datos,
    )


def _subir_en_partes(
    cliente: ClienteGraph, carpeta: Carpeta, nombre: str, datos: bytes
) -> dict:
    """Upload session. Con un extracto no deberia usarse nunca; esta por las dudas."""
    sesion = cliente.json(
        "POST",
        f"/drives/{carpeta.drive_id}/items/{carpeta.item_id}:/{quote(nombre)}"
        ":/createUploadSession",
        json={"item": {"@microsoft.graph.conflictBehavior": "replace", "name": nombre}},
    )
    url = sesion["uploadUrl"]
    total = len(datos)
    resp = None
    for inicio in range(0, total, TAMANIO_TROZO):  # tienen que ir en orden
        fin = min(inicio + TAMANIO_TROZO, total) - 1
        resp = _subir_trozo(url, datos[inicio : fin + 1], inicio, fin, total)
        if resp.status_code not in (200, 201, 202):
            detalle = f"{resp.status_code} {resp.text[:300]}"
            _cancelar_sesion(url)  # el DELETE va despues de leer la respuesta
            raise RuntimeError(f"Fallo el trozo {inicio}-{fin} de {nombre}: {detalle}")
    return resp.json()


def _subir_trozo(
    url: str, trozo: bytes, inicio: int, fin: int, total: int
) -> requests.Response:
    """Sube un trozo, reintentandolo si se corta la red.

    Es para lo que existe partir el archivo: Graph acepta que se reenvie el
    mismo rango mientras la sesion viva, asi que un corte no obliga a resubir
    los trozos que ya entraron.
    """
    ultimo: Exception | None = None
    for intento in range(REINTENTOS):
        try:
            # La uploadUrl ya viene autorizada: mandarle el Bearer la rechaza.
            return requests.put(
                url,
                data=trozo,
                headers={"Content-Range": f"bytes {inicio}-{fin}/{total}"},
                timeout=TIMEOUT,
            )
        except requests.exceptions.RequestException as error:
            ultimo = error
            if intento == REINTENTOS - 1:
                break
            espera = min(2 ** (intento + 1), ESPERA_MAXIMA)
            print(
                f"Se corto la subida del trozo {inicio}-{fin}; "
                f"reintentando en {espera}s...",
                flush=True,
            )
            time.sleep(espera)
    _cancelar_sesion(url)
    raise RuntimeError(_explicar_error_red(ultimo, "PUT", url)) from ultimo


def _cancelar_sesion(url: str) -> None:
    """Borra la upload session para no dejarla colgada hasta que expire.

    Va en su propio try: si la red sigue caida, el error del DELETE taparia el
    del trozo, que es el que explica que paso.
    """
    try:
        requests.delete(url, timeout=(15, 30))
    except requests.exceptions.RequestException:
        pass


# --- volcado ---------------------------------------------------------------
@dataclass
class ResumenVolcado:
    agregados: int
    ya_existentes: int
    url: str = ""


def volcar(
    cliente: ClienteGraph,
    carpeta: Carpeta,
    movimientos: list[Movimiento],
    juego: JuegoDeReglas,
    nombre: str = NOMBRE_POR_DEFECTO,
    respetar_categoria_existente: bool = True,
) -> ResumenVolcado:
    """Baja el extracto que ya esta en la carpeta, mergea, y lo vuelve a subir."""
    existente = buscar(cliente, carpeta, nombre)
    previos: list[Movimiento] = []
    if existente:
        crudo = descargar(cliente, carpeta, existente.id)
        previos = leer_libro_existente(crudo, nombre)

    todos, agregados = combinar(
        previos, movimientos, juego, respetar_categoria_existente
    )
    # Si no hay nada nuevo no se sube nada: no tiene sentido generar una version
    # mas en el historial de SharePoint con el mismo contenido.
    if existente and not agregados:
        return ResumenVolcado(0, len(previos), carpeta.url_web)

    if existente:
        # El libro se regenera entero y se sube con conflictBehavior=replace:
        # de lo que haya del lado de SharePoint solo sobreviven las 6 columnas
        # de 'Cobranzas' (ver la tabla de diferencias del README).
        print(
            f"Aviso: {nombre} se reescribe entero. Las columnas, hojas o "
            "formatos que alguien haya agregado en SharePoint se pierden; "
            "las categorias corregidas a mano se conservan.",
            flush=True,
        )

    subido = subir(
        cliente,
        carpeta,
        nombre,
        armar_libro(todos, juego.categorias_validas),
        etag=existente.etag if existente else None,
    )
    return ResumenVolcado(agregados, len(previos), subido.get("webUrl", ""))

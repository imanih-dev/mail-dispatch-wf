import base64
import os
import re
from typing import List, Dict, Any, Optional

from pydantic import BaseModel, EmailStr, StrictBool, field_validator
from exchangelib import (
    Message as ExMessage,
    HTMLBody,
    Mailbox,
    FileAttachment,
    Account,
    Credentials,
)

class AttachmentManifest(BaseModel):
    content_bytes: str
    filename: str
    cid: Optional[str] = None

    @field_validator("filename")
    def must_have_extension(cls, v: str) -> str:
        if '.' not in v or v.startswith('.') or v.endswith('.'):
            raise ValueError("❌ Filename must have a valid extension")
        return v

    def decoded_content(self) -> bytes:
        try:
            return base64.b64decode(self.content_bytes)
        except Exception:
            raise ValueError("❌ Invalid base64 content in attachment")


class Message(BaseModel):
    account_name: str
    subject: str
    to_recipients: List[EmailStr]
    cc_recipients: Optional[List[EmailStr]] = []
    attachments: Optional[List[Dict[str, Any]]] = []
    html_body: str
    use_signature: Optional[StrictBool] = False


class EmailClient:
    """
    Clase para gestionar la conexión a una cuenta de correo de Exchange
    y enviar mensajes con adjuntos y firmas.
    """

    def __init__(self, email: EmailStr, password: str):
        self._account = self._load_account(email, password)

    def _load_account(self, email: EmailStr, password: str) -> Account:
        """Carga la cuenta de Exchange para el envío de correos."""
        credentials = Credentials(username=str(email), password=password)
        return Account(
            primary_smtp_address=str(email),
            credentials=credentials,
            autodiscover=True,
        )

    def _replace_src_with_cid(self, html: str, res_folder: str, filename: str) -> str:
        """Reemplaza las rutas locales de imágenes con 'cid' para adjuntos en línea."""
        pattern = re.compile(
            rf'(?i)\b((?:src|background)\s*=\s*)([\'"])(?:'
            rf'.*?\b{re.escape(res_folder)}[\\/])?'
            rf'({re.escape(filename)})(?:\?.*?)?([\'"])'
        )

        def repl(m):
            attr_eq, quote_open, _, quote_close = m.groups()
            return f"{attr_eq}{quote_open}cid:{filename}{quote_close}"
        
        return pattern.sub(repl, html)


    def _load_signature(self, signature_key: str):
        """Carga el HTML y los adjuntos de una firma de Outlook."""
        sig_dir = os.path.join(os.environ["APPDATA"], "Microsoft", "Signatures")
        sig_file = os.path.join(sig_dir, f"{signature_key}.htm")
        
        if not os.path.exists(sig_file):
            raise FileNotFoundError(f"Signature '{signature_key}' not found")
        
        try:
            with open(sig_file, encoding="utf-8") as f:
                signature_html = f.read()
        except UnicodeDecodeError:
            with open(sig_file, encoding="latin-1") as f:
                signature_html = f.read()
        
        res_folder_path = None
        res_folder_name = None
        if os.path.isdir(sig_dir):
            candidates = [
                d for d in os.listdir(sig_dir)
                if os.path.isdir(os.path.join(sig_dir, d)) and d.lower().startswith(signature_key.lower())
            ]
            if candidates:
                res_folder_name = sorted(candidates, key=len, reverse=True)[0]
                res_folder_path = os.path.join(sig_dir, res_folder_name)
        
        attachments = []
        if res_folder_path and os.path.exists(res_folder_path):
            for resource_file in os.listdir(res_folder_path):
                if resource_file.lower().endswith((".xml", ".txt", ".htm", ".html", ".thmx")):
                    continue
                
                file_path = os.path.join(res_folder_path, resource_file)
                if not os.path.isfile(file_path):
                    continue
                
                with open(file_path, "rb") as f:
                    content_bytes = base64.b64encode(f.read()).decode("utf-8")
                
                signature_html = self._replace_src_with_cid(signature_html, res_folder_name, resource_file)
                
                attachments.append({
                    "filename": resource_file,
                    "cid": resource_file,
                    "content_bytes": content_bytes,
                })
        
        return signature_html, attachments

    def send_message(self, message: Message, signature_key: Optional[str] = None):
        """Envía un mensaje de correo electrónico con todos sus componentes."""
        msg_html_body: str = message.html_body
        msg_attachments: List[Dict[str, Any]] = message.attachments or []
        
        if signature_key is not None and message.use_signature:
            signature_html, signature_attachments = self._load_signature(signature_key)
            
            if not (signature_html is None or signature_attachments is None):
                msg_html_body = f"{msg_html_body}<br><br>{signature_html}"
                msg_attachments += signature_attachments
        
        email = ExMessage(
            account=self._account,
            folder=self._account.sent,
            subject=message.subject,
            body=HTMLBody(msg_html_body),
            to_recipients=[Mailbox(email_address=addr) for addr in message.to_recipients],
            cc_recipients=[Mailbox(email_address=addr) for addr in (message.cc_recipients or [])],
        )
        
        for att_dict in msg_attachments:
            attachment = AttachmentManifest.model_validate(att_dict)
            
            filename = attachment.filename
            cid = attachment.cid
            content = attachment.decoded_content()
            
            file_attachment = FileAttachment(name=filename, content=content)
            
            if cid:
                valid_extensions = (".jpg", ".jpeg", ".png", ".bmp")
                if filename.lower().endswith(valid_extensions):
                    file_attachment.is_inline = True
                    file_attachment.content_id = cid
                else:
                    raise TypeError(f"❌ File {filename} can not be an inline image")
            
            email.attach(file_attachment)
            
        try:
            email.send()
        except Exception as e:
            raise SystemError(f"❌ Can not send the email: <{e}>")

    @staticmethod
    def list_signatures() -> List[str]:
        """Lista las firmas de Outlook disponibles en el sistema."""
        sig_dir = os.path.join(os.environ["APPDATA"], "Microsoft", "Signatures")
        if not os.path.exists(sig_dir):
            print("No signature directory found.")
            return []
        signatures = [f for f in os.listdir(sig_dir) if f.endswith(".htm")]
        return [os.path.splitext(sig)[0] for sig in signatures]

def render_mdx(template: str, context: Dict[str, Any]) -> str:
    def numtriplet(value: str) -> str:
        try:
            num = float(value)
            return f"{num:,.2f}"
        except (ValueError, TypeError):
            return value

    transformations = {
        'upper': str.upper,
        'lower': str.lower,
        'title': str.title,
        'capitalize': str.capitalize,
        'numtriplet': numtriplet
    }

    unified_pattern = re.compile(
        r'::repeat\s+(\w+)\s*\n(.*?)::endrepeat|{(\w+)\s*\|\s*(\w+)}|{(\w+)}',
        re.DOTALL
    )

    def resolve_path(path: str, ctx: Dict[str, Any]):
        val = ctx
        for part in path.split('.'):
            if isinstance(val, dict):
                val = val.get(part)
            else:
                val = getattr(val, part, None)
            if val is None:
                raise KeyError(f"Path part '{part}' not found in context.")
        return val

    def replace_placeholder(match):
        repeat_match = match.group(1)
        if repeat_match:
            list_name = repeat_match
            block = match.group(2).strip('\n')
            
            try:
                items = resolve_path(list_name, context)
                if not isinstance(items, list):
                    return f"{{ERROR: '{list_name}' is not a list}}"
                
                rendered_items = []
                for item in items:
                    rendered_block = re.sub(
                        r'{(\w+)\s*\|\s*(\w+)}|{(\w+)}',
                        lambda m: process_item_placeholder(m, item),
                        block
                    )
                    rendered_items.append(rendered_block)
                return '\n'.join(rendered_items)
            except (KeyError, AttributeError):
                return f"{{ERROR: List '{list_name}' not found}}"
        
        trans_match_key, trans_match_func, simple_match_key = match.groups()[2:]
        key = trans_match_key if trans_match_key else simple_match_key
        func_name = trans_match_func

        try:
            value = resolve_path(key, context)
            if func_name:
                if func_name in transformations:
                    return transformations[func_name](str(value))
                else:
                    return f"{{ERROR: Func '{func_name}' not found}}"
            else:
                return str(value)
        except KeyError:
            return f"{{ERROR: Key '{key}' not found}}"
        except AttributeError:
            return f"{{ERROR: Invalid path for '{key}'}}"
            
    def process_item_placeholder(match, item_context):
        trans_match_key, trans_match_func, simple_match_key = match.groups()
        key = trans_match_key if trans_match_key else simple_match_key
        func_name = trans_match_func

        try:
            value = resolve_path(key, item_context)
            if func_name:
                if func_name in transformations:
                    return transformations[func_name](str(value))
                else:
                    return f"{{ERROR: Func '{func_name}' not found}}"
            else:
                return str(value)
        except KeyError:
            return f"{{ERROR: Key '{key}' not found in repeat item}}"
        except AttributeError:
            return f"{{ERROR: Invalid path for '{key}' in repeat item}}"

    return unified_pattern.sub(replace_placeholder, template)
# Ejemplo de uso
# CÓDIGO CONGELADO: No debería mostrarse este código
#
# from mi_libreria import EmailClient, Message
#
# email_client = EmailClient(
#     email="tu_correo@ejemplo.com", 
#     password="tu_password"
# )
#
# lista_de_firmas = EmailClient.list_signatures()
#
# mensaje = Message(
#     account_name="Mi empresa",
#     subject="Prueba de envío",
#     to_recipients=["destinatario@ejemplo.com"],
#     html_body="<h1>Hola</h1><p>Este es un mensaje de prueba.</p>",
#     use_signature=True,
# )
#
# email_client.send_message(mensaje, signature_key="MiFirma")
import base64
from typing import List
from core.models import EmailItem

DEMO_LOGO_SVG = """<svg xmlns="http://www.w3.org/2000/svg" width="260" height="96" viewBox="0 0 260 96">
 <rect width="260" height="96" fill="#fff"/>
 <polygon points="112,52 134,14 156,52 146,52 134,31 122,52" fill="#2b2b2b"/>
 <polygon points="134,14 156,52 147,52 130,21" fill="#e3262f"/>
 <text x="134" y="72" font-family="Arial" font-weight="bold" font-size="13" text-anchor="middle" fill="#333">AMERICAN TOWER</text>
 <text x="134" y="84" font-family="Arial" font-size="7" letter-spacing="2" text-anchor="middle" fill="#999">DO BRASIL</text>
</svg>"""

DEMO_LOGO_URI = "data:image/svg+xml;base64," + base64.b64encode(DEMO_LOGO_SVG.encode("utf-8")).decode("ascii")

DEMO_CLIENTS = [
    {
        "to": "administrativo@3rjjf.com.br",
        "nome": "PRIME INTERNET",
        "list": "MTR-CORP-5562 &nbsp;&nbsp;&nbsp; 3. REGISTRO DE IMOVEIS DE JUIZ DE FORA-MG - 450671 - Prime Internet",
        "sent": "23/10/2026 09:12",
    },
    {
        "to": "noc@vianettelecom.com.br",
        "nome": "VIANET TELECOM",
        "list": "MTR-CORP-5571 &nbsp;&nbsp;&nbsp; VIANET POP CENTRO JF - 450702 - ViaNet Telecom",
        "sent": "23/10/2026 09:14",
    },
    {
        "to": "ti@hospitalmonteserrat.org.br",
        "nome": "HOSPITAL MONTE SERRAT",
        "list": "MTR-CORP-5588 &nbsp;&nbsp;&nbsp; HOSPITAL MONTE SERRAT - BLOCO B - 450733 - Link Dedicado",
        "sent": "23/10/2026 09:15",
    },
    {
        "to": "infra@pjf.mg.gov.br",
        "nome": "PREFEITURA DE JUIZ DE FORA",
        "list": "MTR-CORP-5590 &nbsp;&nbsp;&nbsp; PREFEITURA JF - SEDE ADMINISTRATIVA - 450741 - MPLS<br>MTR-CORP-5591 &nbsp;&nbsp;&nbsp; PREFEITURA JF - ANEXO II - 450742 - MPLS",
        "sent": "23/10/2026 09:17",
    },
    {
        "to": "operacoes@conectmais.net.br",
        "nome": "CONECT MAIS",
        "list": "MTR-CORP-5603 &nbsp;&nbsp;&nbsp; CONECT MAIS NODE ALTO DOS PASSOS - 450760 - Trânsito IP",
        "sent": "23/10/2026 09:18",
    },
    {
        "to": "suporte@colegiosantamaria.edu.br",
        "nome": "COLÉGIO SANTA MARIA",
        "list": "MTR-CORP-5611 &nbsp;&nbsp;&nbsp; COLEGIO SANTA MARIA - UNIDADE CENTRO - 450777 - Internet Dedicada",
        "sent": "23/10/2026 09:20",
    },
]


def build_demo_html(client_row: dict) -> str:
    return f"""<!DOCTYPE html><html><head><meta charset="utf-8"><style>
body{{font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#222;background:#fff;margin:18px 22px}}
p{{margin:0 0 4px}}.hdr{{text-align:right}}.hdr img{{border:1px solid #ddd}}
.hdr .t{{font-family:Arial;font-size:13pt;color:#333;text-decoration:underline;margin-top:4px}}
</style></head><body>
<div class="hdr"><img src="cid:logo_atc" width="260"><div class="t">Informativo de Interrupção</div></div>
<p>&nbsp;</p><p>Prezado cliente,</p><p>&nbsp;</p>
<p>A American Tower, através de seu time de engenharia e operações, vem por meio desse comunicado informar que serão realizadas atividades em nossa rede que ocasionarão interrupção nos serviços prestados durante a janela de execução.</p>
<p>&nbsp;</p>
<p><b><u>Seguem abaixo dados relevantes em relação à atividade.</u></b></p>
<p>&nbsp;</p>
<p><b>Registro da atividade:</b> 655957</p>
<p><b>Localidade(s) de execução:</b> Juiz de Fora / MG</p>
<p><b>Data/Hora de início:</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 04/11/2026 – 00h00</p>
<p><b>Data/Hora de término:</b>&nbsp;&nbsp;&nbsp; 04/11/2026 – 06h00</p>
<p><b>Descrição/Justificativa:</b> Manutenção Corretiva - Substituição e padronização da caixa de emendas CE-JFA-216B.</p>
<p><b>Listagem dos equipamentos, clientes e/ou nodes afetados do cliente:</b></p>
<p style="font-size:9pt;margin-left:6px">{client_row["list"]}</p>
<p>&nbsp;</p>
<p>Aproveitamos a oportunidade para reiterar nosso compromisso com a melhoria contínua da qualidade e disponibilidade de nossos serviços. A American Tower, através de sua equipe de atendimento e operações, segue à disposição para quaisquer esclarecimentos que se façam necessários.</p>
<p>&nbsp;</p><p>Atenciosamente,</p>
<p><b>Centro de Operações de Rede – American Tower do Brasil</b></p>
</body></html>"""


def get_demo_emails() -> List[EmailItem]:
    out: List[EmailItem] = []
    for idx, c in enumerate(DEMO_CLIENTS):
        html = build_demo_html(c)
        out.append(
            EmailItem(
                id=f"e{idx}",
                entry_id="",
                store_id="",
                subject="INFORMATIVO DE ATIVIDADE PROGRAMADA - AMERICAN TOWER",
                to=c["to"],
                cc="",
                bcc="",
                sent_on=c["sent"],
                sender_name="sir.fibra",
                html_body=html,
                text_body="",
                inline_images={"logo_atc": DEMO_LOGO_URI},
                attachments_count=0,
            )
        )
    return out

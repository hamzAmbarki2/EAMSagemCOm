"""Writes presentation/use_case_diagram.drawio (open at app.diagrams.net).
Same content as the TikZ use case diagram in rapport/chapters/chap3_conception.tex."""
import os
from xml.sax.saxutils import quoteattr
S, H = 45, 22.5
cells, n = [], [1]
def cid():
    n[0] += 1; return f"c{n[0]}"
def px(x, y): return x * S, (H - y) * S
def add(xml): cells.append(xml)
add(f'<mxCell id="sys" value="EAM APP" style="rounded=0;whiteSpace=wrap;verticalAlign=top;fontStyle=1;fontSize=14;fillColor=none;strokeWidth=2;" vertex="1" parent="1"><mxGeometry x="{3.5*S}" y="{(H-22.5)*S}" width="{15.5*S}" height="{24.7*S}" as="geometry"/></mxCell>')
ids = {}
def actor(key, label, x, y):
    i = cid(); ids[key] = i; X, Y = px(x, y)
    add(f'<mxCell id="{i}" value={quoteattr(label)} style="shape=umlActor;verticalLabelPosition=bottom;verticalAlign=top;html=1;" vertex="1" parent="1"><mxGeometry x="{X-15}" y="{Y-30}" width="30" height="60" as="geometry"/></mxCell>')
def uc(key, label, x, y, w=140, h=62, small=False):
    i = cid(); ids[key] = i; X, Y = px(x, y)
    fs = 10 if small else 12
    add(f'<mxCell id="{i}" value={quoteattr(label)} style="ellipse;whiteSpace=wrap;html=1;fontSize={fs};" vertex="1" parent="1"><mxGeometry x="{X-w/2}" y="{Y-h/2}" width="{w}" height="{h}" as="geometry"/></mxCell>')
def edge(a, b, label="", dashed=False):
    i = cid()
    st = "endArrow=open;html=1;fontSize=9;" + ("dashed=1;endArrow=open;" if dashed else "endArrow=none;")
    add(f'<mxCell id="{i}" value={quoteattr(label)} style="{st}" edge="1" parent="1" source="{ids[a]}" target="{ids[b]}"><mxGeometry relative="1" as="geometry"/></mxCell>')
for k, l, y in [("chefop", "Chefop", 20.8), ("thech", "Thech", 13.5), ("cheftech", "ChefTech", 6.5), ("admin", "Admin", 0.5)]:
    actor(k, l, 1.5, y)
for k, l, y in [("eth", "manage labels", 21.5), ("stat", "view machine status", 19.6), ("bt", "manage intervention orders", 17.6),
                ("plan", "view schedule", 15.6), ("rap", "view reports", 11.5), ("arch", "manage archive", 9.7),
                ("ot", "manage work order tasks", 7.8), ("alerte", "manage alerts", 5.9), ("gplan", "manage schedule", 2.3),
                ("roles", "manage roles/users", 0.5)]:
    uc(k, l, 8.5, y)
for k, l, x, y in [("jr", "by day", 5.5, 13.4), ("sem", "by week", 8.5, 13.4), ("mois", "by month", 11.5, 13.4), ("prio", "by priority", 5.5, 4.0)]:
    uc(k, l, x, y, w=100, h=44, small=True)
uc("auth", "authentication", 16.0, 11.0, w=150, h=150)
for a, b in [("chefop", "eth"), ("chefop", "stat"), ("thech", "bt"), ("thech", "plan"), ("thech", "rap"), ("thech", "arch"),
             ("cheftech", "ot"), ("cheftech", "alerte"), ("admin", "gplan"), ("admin", "roles")]:
    edge(a, b)
for a, b in [("plan", "jr"), ("plan", "sem"), ("plan", "mois"), ("alerte", "prio")]:
    edge(a, b, "<<extends>>", True)
for a in ["eth", "stat", "bt", "plan", "rap", "arch", "ot", "alerte", "gplan", "roles"]:
    edge(a, "auth", "<<includes>>" if a == "plan" else "", True)
out = ('<mxfile><diagram name="Use case diagram"><mxGraphModel dx="1200" dy="900" grid="1" gridSize="10" page="1" pageWidth="1600" pageHeight="1200">'
       '<root><mxCell id="0"/><mxCell id="1" parent="0"/>' + "".join(cells) + '</root></mxGraphModel></diagram></mxfile>')
p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "use_case_diagram.drawio")
open(p, "w", encoding="utf-8").write(out)
print("saved", p)

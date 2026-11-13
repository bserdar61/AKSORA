Attribute VB_Name = "modAKSORA_PANO_SABLON"
Option Explicit

' OTOMATIK URETILIR - ELLE DUZENLEMEYIN.
' Kaynak: gelistirme/pano/pano_sablon.html  |  Uretim: python3 tools/aksora_vba.py sablon ...

Public Function AKSORA_PANO_SABLON() As String
    AKSORA_PANO_SABLON = AKSORA_PANO_SABLON & PS_01() & PS_02() & PS_03() & PS_04() & PS_05()
End Function

Private Function PS_01() As String
    Dim s As String
    s = s & "<!doctype html>" & vbLf
    s = s & "<html lang=""tr"">" & vbLf
    s = s & "<head>" & vbLf
    s = s & "<meta charset=""utf-8"">" & vbLf
    s = s & "<meta name=""viewport"" content=""width=device-width, initial-scale=1"">" & vbLf
    s = s & "<title>AKSORA Yönetici Panosu</title>" & vbLf
    s = s & "<style>" & vbLf
    s = s & ":root{color-scheme:light;--bg:#f3f6fa;--s1:#fcfcfb;--ink:#0b0b0b;--ink2:#52514e;--ink3:#8a8984;--line:#e3e6ea;--brand:#0c2d4f;--brand2:#17426b;--lime:#c6f061;--c1:#2a78d6;--c1s:#cde2fb;--good:#0ca30c;--warn:#fab219;--ser:#ec835a;--crit:#d03b3b;--chip:#eef2f7}" & vbLf
    s = s & "@media (prefers-color-scheme:dark){:root{color-scheme:dark;--bg:#111110;--s1:#1a1a19;--ink:#f0efec;--ink2:#c3c2b7;--ink3:#8f8e86;--line:#30302d;--brand:#0c2d4f;--brand2:#17426b;--c1:#3987e5;--c1s:#184f95;--chip:#262624}}" & vbLf
    s = s & "*{box-sizing:border-box}" & vbLf
    s = s & "body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 ""Segoe UI"",system-ui,-apple-system,Roboto,Arial,sans-serif}" & vbLf
    s = s & "header{background:linear-gradient(120deg,var(--brand),var(--brand2));color:#fff;padding:18px 16px}" & vbLf
    s = s & ".wrap{max-width:1180px;margin:0 auto}" & vbLf
    s = s & "h1{margin:0;font-size:22px;letter-spacing:.3px}" & vbLf
    s = s & ".sub{margin-top:4px;color:var(--lime);font-size:13px;font-weight:600}" & vbLf
    s = s & ".badge{display:inline-block;margin-left:8px;padding:1px 8px;border-radius:10px;background:#fab219;color:#0b0b0b;font-size:11px;font-weight:700;vertical-align:middle}" & vbLf
    s = s & "main{padding:16px}" & vbLf
    s = s & ".grid{display:grid;gap:12px}" & vbLf
    s = s & ".tiles{grid-template-columns:repeat(4,minmax(0,1fr))}" & vbLf
    s = s & ".two{grid-template-columns:minmax(0,3fr) minmax(0,2fr)}" & vbLf
    s = s & ".half{grid-template-columns:repeat(2,minmax(0,1fr))}" & vbLf
    s = s & "@media (max-width:900px){.tiles{grid-template-columns:repeat(2,minmax(0,1fr))}.two,.half{grid-template-columns:minmax(0,1fr)}}" & vbLf
    s = s & "@media (max-width:460px){.tiles{grid-template-columns:minmax(0,1fr)}}" & vbLf
    s = s & ".card{background:var(--s1);border:1px solid var(--line);border-radius:12px;padding:14px;min-width:0}" & vbLf
    s = s & ".card h2{margin:0 0 10px;font-size:13px;color:var(--ink2);font-weight:700;text-transform:uppercase;letter-spacing:.5px}" & vbLf
    s = s & ".sec{margin-top:12px}" & vbLf
    s = s & ".tl{font-size:12px;color:var(--ink2);font-weight:600}" & vbLf
    s = s & ".tv{font-size:26px;font-weight:700;margin:2px 0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}" & vbLf
    s = s & ".td{font-size:12px;color:var(--ink2);display:flex;gap:6px;align-items:center;flex-wrap:wrap;min-height:20px}" & vbLf
    s = s & ".pill{display:inline-flex;align-items:center;gap:4px;padding:1px 7px;border-radius:10px;background:var(--chip);color:var(--ink);font-size:11px;font-weight:600}" & vbLf
    s = s & ".dot{width:8px;height:8px;border-radius:50%;display:inline-block}" & vbLf
    s = s & ".spark{display:block;width:100%;height:34px;margin-top:6px}" & vbLf
    s = s & "ol.jobs{margin:0;padding:0;list-style:none}" & vbLf
    s = s & "ol.jobs li{display:grid;grid-template-columns:28px minmax(0,1fr) auto;gap:10px;padding:9px 0;border-top:1px solid var(--line);align-items:start}" & vbLf
    s = s & "ol.jobs li:first-child{border-top:0}" & vbLf
    s = s & ".no{width:26px;height:26px;border-radius:50%;background:var(--brand);color:#fff;font-weight:700;display:flex;align-items:center;justify-content:center;font-size:13px}" & vbLf
    s = s & ".amt{font-weight:700;white-space:nowrap}" & vbLf
    s = s & ".where{font-size:12px;color:var(--ink3)}" & vbLf
    s = s & ".seg{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px}" & vbLf
    s = s & ".seg button{border:1px solid var(--line);background:var(--s1);color:var(--ink2);border-radius:16px;padding:4px 10px;font:inherit;font-size:12px;cursor:pointer}" & vbLf
    PS_01 = s
End Function

Private Function PS_02() As String
    Dim s As String
    s = s & ".seg button[aria-pressed=true]{background:var(--brand);border-color:var(--brand);color:#fff}" & vbLf
    s = s & ".chart{position:relative}" & vbLf
    s = s & ".chart svg{display:block;width:100%;height:230px;touch-action:pan-y}" & vbLf
    s = s & ".axis{fill:var(--ink3);font-size:11px}" & vbLf
    s = s & ".gridl{stroke:var(--line);stroke-width:1}" & vbLf
    s = s & ".tip{position:absolute;pointer-events:none;background:var(--ink);color:var(--s1);padding:5px 8px;border-radius:6px;font-size:12px;white-space:nowrap;transform:translate(-50%,-120%);display:none;z-index:5}" & vbLf
    s = s & ".bars{display:grid;grid-template-columns:auto minmax(0,1fr) auto;gap:8px 10px;align-items:center}" & vbLf
    s = s & ".bl{font-weight:600;white-space:nowrap}" & vbLf
    s = s & ".bt{height:14px;position:relative}" & vbLf
    s = s & ".bf{height:14px;background:var(--c1);border-radius:0 4px 4px 0;min-width:2px}" & vbLf
    s = s & ".bv{font-variant-numeric:tabular-nums;white-space:nowrap;color:var(--ink2);font-size:13px}" & vbLf
    s = s & ".note{font-size:12px;color:var(--ink3);margin-top:8px}" & vbLf
    s = s & "ul.al{margin:0;padding:0;list-style:none}" & vbLf
    s = s & "ul.al li{display:grid;grid-template-columns:auto minmax(0,1fr);gap:10px;padding:8px 0;border-top:1px solid var(--line)}" & vbLf
    s = s & "ul.al li:first-child{border-top:0}" & vbLf
    s = s & "ul.al .pill{align-self:start;white-space:nowrap}" & vbLf
    s = s & ".tbl{overflow-x:auto}" & vbLf
    s = s & "table{border-collapse:collapse;width:100%;font-size:13px}" & vbLf
    s = s & "th,td{padding:7px 8px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}" & vbLf
    s = s & "th:first-child,td:first-child{text-align:left}" & vbLf
    s = s & "th{color:var(--ink2);font-size:12px}" & vbLf
    s = s & "footer{color:var(--ink3);font-size:12px;padding:6px 16px 24px}" & vbLf
    s = s & ".empty{color:var(--ink3);font-size:13px;padding:8px 0}" & vbLf
    s = s & "@media print{header{-webkit-print-color-adjust:exact;print-color-adjust:exact}.seg{display:none}body{background:#fff}}" & vbLf
    s = s & "</style>" & vbLf
    s = s & "</head>" & vbLf
    s = s & "<body>" & vbLf
    s = s & "<header><div class=""wrap""><h1>AKSORA &nbsp;|&nbsp; Yönetici Panosu <span id=""ornek""></span></h1><div class=""sub"" id=""sub""></div></div></header>" & vbLf
    s = s & "<main><div class=""wrap"">" & vbLf
    s = s & "<div class=""grid tiles"" id=""tiles""></div>" & vbLf
    s = s & "<div class=""grid two sec"">" & vbLf
    s = s & "<div class=""card""><h2>Bugün yapılacak işler</h2><ol class=""jobs"" id=""jobs""></ol></div>" & vbLf
    s = s & "<div class=""card""><h2>Öncelikli uyarılar</h2><ul class=""al"" id=""alerts""></ul></div>" & vbLf
    s = s & "</div>" & vbLf
    s = s & "<div class=""card sec""><h2 id=""trh"">Eğilim</h2><div class=""seg"" id=""seg""></div><div class=""chart"" id=""trend""><div class=""tip"" id=""tip""></div></div><div class=""note"" id=""trn""></div></div>" & vbLf
    s = s & "<div class=""grid half sec"">" & vbLf
    s = s & "<div class=""card""><h2>İllere göre tekil stok değeri</h2><div class=""bars"" id=""bil""></div></div>" & vbLf
    s = s & "<div class=""card""><h2>İllere göre atıl stok oranı</h2><div class=""bars"" id=""bat""></div><div class=""note"" id=""batn""></div></div>" & vbLf
    s = s & "</div>" & vbLf
    s = s & "<div class=""grid half sec"">" & vbLf
    PS_02 = s
End Function

Private Function PS_03() As String
    Dim s As String
    s = s & "<div class=""card""><h2>Birimlere göre stok değeri</h2><div class=""bars"" id=""bbr""></div><div class=""note"">Birimler ortak depoları paylaşır; değerler toplanmaz.</div></div>" & vbLf
    s = s & "<div class=""card""><h2>Gösterge tablosu</h2><div class=""tbl""><table id=""tb""></table></div></div>" & vbLf
    s = s & "</div>" & vbLf
    s = s & "</div></main>" & vbLf
    s = s & "<footer><div class=""wrap"" id=""foot""></div></footer>" & vbLf
    s = s & "<div class=""tip"" id=""btip"" style=""position:fixed""></div>" & vbLf
    s = s & "<script>" & vbLf
    s = s & "var D=/*VERI*/null/*VERI*/;" & vbLf
    s = s & "(function(){" & vbLf
    s = s & """use strict"";" & vbLf
    s = s & "var $=function(i){return document.getElementById(i)};" & vbLf
    s = s & "var nf=function(x,d){return Number(x).toLocaleString(""tr-TR"",{minimumFractionDigits:d,maximumFractionDigits:d})};" & vbLf
    s = s & "function esc(s){return String(s==null?"""":s).replace(/[&<>""]/g,function(c){return{""&"":""&amp;"",""<"":""&lt;"","">"":""&gt;"",'""':""&quot;""}[c]})}" & vbLf
    s = s & "function tl(x){var a=Math.abs(x);if(a>=1e9)return nf(x/1e9,2)+"" mr TL"";if(a>=1e6)return nf(x/1e6,1)+"" mn TL"";if(a>=1e3)return nf(x/1e3,1)+"" bin TL"";return nf(x,2)+"" TL""}" & vbLf
    s = s & "function fmt(v,t){if(v==null)return ""veri yok"";if(t==1)return nf(v,0)+"" / 100"";if(t==2)return tl(v);if(t==3)return ""%""+nf(v*100,1);return nf(v,0)}" & vbLf
    s = s & "function dfmt(d,t){var a=Math.abs(d);if(t==1)return nf(a,0)+"" puan"";if(t==2)return tl(a);if(t==3)return nf(a*100,1)+"" puan"";return nf(a,0)}" & vbLf
    s = s & "function delta(k,o){" & vbLf
    s = s & " if(k.bugun==null||o==null)return{txt:""karşılaştırma yok"",st:0,lab:""""};" & vbLf
    s = s & " var d=k.bugun-o;if(Math.abs(d)<5e-5)return{txt:""değişmedi"",st:0,lab:""Aynı""};" & vbLf
    s = s & " var st=k.yon==0?0:((d>0)==(k.yon>0)?1:-1);" & vbLf
    s = s & " return{txt:(d>0?""\u25B2 "":""\u25BC "")+dfmt(d,k.tur),st:st,lab:st>0?""İyileşti"":st<0?""Kötüleşti"":(d>0?""Arttı"":""Azaldı"")};" & vbLf
    s = s & "}" & vbLf
    s = s & "function pill(dl){if(!dl.lab)return """";var c=dl.st>0?""var(--good)"":dl.st<0?""var(--crit)"":""var(--ink3)"";var ic=dl.st>0?""\u2713"":dl.st<0?""!"":""•"";return '<span class=""pill""><span class=""dot"" style=""background:'+c+'""></span>'+ic+"" ""+dl.lab+""</span>""}" & vbLf
    s = s & "function pts(a){var o=[];for(var i=0;i<a.length;i++)if(a[i]!=null)o.push([i,a[i]]);return o}" & vbLf
    s = s & "function spark(a){" & vbLf
    s = s & " var p=pts(a);if(p.length<2)return """";" & vbLf
    s = s & " var mn=Infinity,mx=-Infinity;p.forEach(function(q){mn=Math.min(mn,q[1]);mx=Math.max(mx,q[1])});if(mx-mn<1e-9){mx=mn+1;mn=mn-1}" & vbLf
    s = s & " var w=200,h=34,x0=p[0][0],x1=p[p.length-1][0];" & vbLf
    s = s & " var d=p.map(function(q,i){var x=4+(w-8)*(q[0]-x0)/Math.max(1,x1-x0),y=h-5-(h-10)*(q[1]-mn)/(mx-mn);return(i?""L"":""M"")+x.toFixed(1)+"" ""+y.toFixed(1)}).join("""");" & vbLf
    s = s & " var l=d.split(""L"").pop().split("" "");" & vbLf
    s = s & " return '<svg class=""spark"" viewBox=""0 0 '+w+"" ""+h+'"" preserveAspectRatio=""none"" aria-hidden=""true""><path d=""'+d+'"" fill=""none"" stroke=""var(--c1)"" stroke-width=""2"" vector-effect=""non-scaling-stroke"" stroke-linejoin=""round""/><circle cx=""'+l[0].replace(""M"","""")+'"" cy=""'+l[1]+'"" r=""3"" fill=""var(--c1)""/></svg>';" & vbLf
    s = s & "}" & vbLf
    s = s & "if(!D){document.body.innerHTML='<p style=""padding:24px"">Veri bulunamadı.</p>';return}" & vbLf
    s = s & "$(""ornek"").innerHTML=D.ornek?'<span class=""badge"">ÖRNEK VERİ</span>':"""";" & vbLf
    s = s & "$(""sub"").textContent=""Çoruh EDAŞ  •  ""+D.il+""  •  ""+D.tarih+(D.onceki?""  •  önceki kayıt: ""+D.onceki:"""");" & vbLf
    s = s & "// --- gostergeler (ust kartlar: ilk dort) ---" & vbLf
    s = s & "var K=D.kpi||[];" & vbLf
    s = s & "$(""tiles"").innerHTML=[0,1,2,4].map(function(i){var k=K[i];if(!k)return """";var dl=delta(k,k.onceki);" & vbLf
    s = s & " return '<div class=""card""><div class=""tl"">'+esc(k.ad)+'</div><div class=""tv"">'+fmt(k.bugun,k.tur)+'</div><div class=""td""><span>'+dl.txt+""</span>""+pill(dl)+""</div>""+spark(k.seri||[])+""</div>""}).join("""");" & vbLf
    s = s & "// --- isler ---" & vbLf
    PS_03 = s
End Function

Private Function PS_04() As String
    Dim s As String
    s = s & "var J=D.isler||[];" & vbLf
    s = s & "$(""jobs"").innerHTML=J.length?J.map(function(j,i){return '<li><span class=""no"">'+(i+1)+'</span><div><div>'+esc(j.is)+'</div><div class=""where"">'+esc(j.nerede)+'</div></div><div class=""amt"">'+(j.tutar?tl(j.tutar):"""")+""</div></li>""}).join(""""):'<li class=""empty"" style=""display:block"">Bekleyen önemli iş yok.</li>';" & vbLf
    s = s & "// --- uyarilar ---" & vbLf
    s = s & "var U=D.uyarilar||[];" & vbLf
    s = s & "$(""alerts"").innerHTML=U.length?U.map(function(u){var kr=u.sev==""KRİTİK"";return '<li><span class=""pill""><span class=""dot"" style=""background:'+(kr?""var(--crit)"":""var(--warn)"")+'""></span>'+(kr?""! Kritik"":""\u25B3 Uyarı"")+""</span><div>""+esc(u.metin)+""</div></li>""}).join(""""):'<li class=""empty"" style=""display:block"">Kritik veya dikkat gerektiren uyarı yok.</li>';" & vbLf
    s = s & "// --- egilim ---" & vbLf
    s = s & "var T=D.seriTarih||[],sel=0,tip=$(""tip""),box=$(""trend"");" & vbLf
    s = s & "function trend(){" & vbLf
    s = s & " var k=K[sel],p=pts(k.seri||[]);" & vbLf
    s = s & " $(""trh"").textContent=""Eğilim  •  ""+k.ad;" & vbLf
    s = s & " var old=box.querySelector(""svg"");if(old)box.removeChild(old);" & vbLf
    s = s & " if(p.length<2){$(""trn"").textContent=""Eğilim için en az iki günlük kayıt gerekir. Yönetim Merkezi her yenilendiğinde günde bir kayıt tutulur."";return}" & vbLf
    s = s & " $(""trn"").textContent=p.length+"" kayıt  •  ""+T[p[0][0]]+"" – ""+T[p[p.length-1][0]];" & vbLf
    s = s & " var W=box.clientWidth||600,H=230,L=64,R=18,Tp=12,B=26;" & vbLf
    s = s & " var mn=Infinity,mx=-Infinity;p.forEach(function(q){mn=Math.min(mn,q[1]);mx=Math.max(mx,q[1])});" & vbLf
    s = s & " var pad=(mx-mn)*0.12;if(pad<1e-9)pad=Math.abs(mx)*0.05||1;mn-=pad;mx+=pad;if(k.tur!=2&&mn<0&&p.every(function(q){return q[1]>=0}))mn=0;" & vbLf
    s = s & " var n=T.length,X=function(i){return L+(W-L-R)*(n>1?i/(n-1):0.5)},Y=function(v){return Tp+(H-Tp-B)*(1-(v-mn)/(mx-mn))};" & vbLf
    s = s & " var s='<svg viewBox=""0 0 '+W+"" ""+H+'"" role=""img"" aria-label=""'+esc(k.ad)+' eğilimi"">';" & vbLf
    s = s & " for(var g=0;g<=3;g++){var v=mn+(mx-mn)*g/3,y=Y(v);s+='<line class=""gridl"" x1=""'+L+'"" x2=""'+(W-R)+'"" y1=""'+y+'"" y2=""'+y+'""/><text class=""axis"" x=""'+(L-8)+'"" y=""'+(y+4)+'"" text-anchor=""end"">'+(k.tur==2?tl(v).replace("" TL"",""""):k.tur==3?""%""+nf(v*100,1):nf(v,0))+""</text>""}" & vbLf
    s = s & " s+='<text class=""axis"" x=""'+X(p[0][0])+'"" y=""'+(H-6)+'"" text-anchor=""start"">'+T[p[0][0]]+'</text><text class=""axis"" x=""'+X(p[p.length-1][0])+'"" y=""'+(H-6)+'"" text-anchor=""end"">'+T[p[p.length-1][0]]+""</text>"";" & vbLf
    s = s & " s+='<path d=""'+p.map(function(q,i){return(i?""L"":""M"")+X(q[0]).toFixed(1)+"" ""+Y(q[1]).toFixed(1)}).join("""")+'"" fill=""none"" stroke=""var(--c1)"" stroke-width=""2"" stroke-linejoin=""round"" stroke-linecap=""round""/>';" & vbLf
    s = s & " var lp=p[p.length-1];s+='<circle cx=""'+X(lp[0])+'"" cy=""'+Y(lp[1])+'"" r=""4.5"" fill=""var(--c1)"" stroke=""var(--s1)"" stroke-width=""2""/>';" & vbLf
    s = s & " s+='<line id=""ch"" class=""gridl"" y1=""'+Tp+'"" y2=""'+(H-B)+'"" style=""display:none;stroke:var(--ink3)""/><circle id=""cd"" r=""4.5"" fill=""var(--c1)"" stroke=""var(--s1)"" stroke-width=""2"" style=""display:none""/></svg>';" & vbLf
    s = s & " box.insertAdjacentHTML(""afterbegin"",s);" & vbLf
    s = s & " var svg=box.querySelector(""svg""),ch=svg.querySelector(""#ch""),cd=svg.querySelector(""#cd"");" & vbLf
    s = s & " function mv(e){var r=svg.getBoundingClientRect(),cx=(e.touches?e.touches[0].clientX:e.clientX)-r.left,sx=cx*W/r.width,b=p[0],bd=1e9;" & vbLf
    s = s & "  p.forEach(function(q){var dd=Math.abs(X(q[0])-sx);if(dd<bd){bd=dd;b=q}});" & vbLf
    s = s & "  var x=X(b[0]),y=Y(b[1]);ch.setAttribute(""x1"",x);ch.setAttribute(""x2"",x);ch.style.display="""";cd.setAttribute(""cx"",x);cd.setAttribute(""cy"",y);cd.style.display="""";" & vbLf
    s = s & "  tip.textContent=T[b[0]]+""  •  ""+fmt(b[1],k.tur);tip.style.display=""block"";" & vbLf
    s = s & "  var px=x*r.width/W;tip.style.left=Math.max(70,Math.min(r.width-70,px))+""px"";tip.style.top=(y*r.height/H)+""px""}" & vbLf
    s = s & " function out(){ch.style.display=""none"";cd.style.display=""none"";tip.style.display=""none""}" & vbLf
    s = s & " svg.addEventListener(""mousemove"",mv);svg.addEventListener(""touchstart"",mv,{passive:true});svg.addEventListener(""touchmove"",mv,{passive:true});svg.addEventListener(""mouseleave"",out);svg.addEventListener(""touchend"",out);" & vbLf
    s = s & "}" & vbLf
    s = s & "$(""seg"").innerHTML=K.map(function(k,i){return '<button type=""button"" data-i=""'+i+'"" aria-pressed=""'+(i==0)+'"">'+esc(k.ad)+""</button>""}).join("""");" & vbLf
    s = s & "$(""seg"").addEventListener(""click"",function(e){var b=e.target.closest(""button"");if(!b)return;sel=+b.getAttribute(""data-i"");[].forEach.call($(""seg"").children,function(c){c.setAttribute(""aria-pressed"",c===b)});trend()});" & vbLf
    s = s & "if(K.length)trend();" & vbLf
    s = s & "var rz;window.addEventListener(""resize"",function(){clearTimeout(rz);rz=setTimeout(function(){if(K.length)trend()},150)});" & vbLf
    s = s & "// --- cubuklar ---" & vbLf
    s = s & "var bt=$(""btip"");" & vbLf
    s = s & "function bars(id,rows,f){" & vbLf
    PS_04 = s
End Function

Private Function PS_05() As String
    Dim s As String
    s = s & " var mx=0;rows.forEach(function(r){if(r.v!=null)mx=Math.max(mx,r.v)});" & vbLf
    s = s & " $(id).innerHTML=rows.length?rows.map(function(r){return '<div class=""bl"">'+esc(r.ad)+'</div><div class=""bt"" data-t=""'+esc(r.ad+""  •  ""+(r.v==null?""veri yok"":f(r.v))+(r.x?""  •  ""+r.x:""""))+'""><div class=""bf"" style=""width:'+(r.v==null||mx<=0?0:Math.max(0.5,100*r.v/mx))+'%""></div></div><div class=""bv"">'+(r.v==null?""–"":f(r.v))+""</div>""}).join(""""):'<div class=""empty"">Veri yok.</div>';" & vbLf
    s = s & " [].forEach.call($(id).querySelectorAll("".bt""),function(el){" & vbLf
    s = s & "  el.addEventListener(""mousemove"",function(e){bt.textContent=el.getAttribute(""data-t"");bt.style.display=""block"";bt.style.left=e.clientX+""px"";bt.style.top=e.clientY+""px""});" & vbLf
    s = s & "  el.addEventListener(""mouseleave"",function(){bt.style.display=""none""})});" & vbLf
    s = s & "}" & vbLf
    s = s & "var I=D.iller||[],top=0;I.forEach(function(r){top+=r.tekil||0});" & vbLf
    s = s & "bars(""bil"",I.map(function(r){return{ad:r.ad,v:r.tekil,x:top>0?""pay %""+nf(100*r.tekil/top,1):""""}}),tl);" & vbLf
    s = s & "if(D.atilVar){bars(""bat"",I.map(function(r){return{ad:r.ad,v:r.kapsam>0?r.atil/r.kapsam:null,x:""atıl ""+tl(r.atil||0)}}),function(v){return ""%""+nf(v*100,1)});$(""batn"").textContent=""Atıl: 24 aydır girişi olmayan stok. Oran = atıl değer / il stok değeri.""}" & vbLf
    s = s & "else{$(""bat"").innerHTML='<div class=""empty"">Atıl stok için Malzeme > ATIL STOK ekranında SAP okuması (MB51) gerekir.</div>'}" & vbLf
    s = s & "bars(""bbr"",(D.birimler||[]).map(function(r){return{ad:r.ad,v:r.deger}}),tl);" & vbLf
    s = s & "// --- tablo ---" & vbLf
    s = s & "$(""tb"").innerHTML=""<tr><th>Gösterge</th><th>Bugün</th><th>Önceki</th><th>Değişim</th><th>30 gün önce</th></tr>""+K.map(function(k){var dl=delta(k,k.onceki);" & vbLf
    s = s & " return ""<tr><td>""+esc(k.ad)+""</td><td><b>""+fmt(k.bugun,k.tur)+""</b></td><td>""+(k.onceki==null?""–"":fmt(k.onceki,k.tur))+""</td><td>""+(k.onceki==null?""–"":dl.txt)+""</td><td>""+(k.otuz==null?""–"":fmt(k.otuz,k.tur))+""</td></tr>""}).join("""");" & vbLf
    s = s & "$(""foot"").textContent=""AKSORA Yönetim Merkezi verilerinden otomatik üretilmiştir  •  Hazırlayan: ""+(D.hazirlayan||""-"")+""  •  Bu dosya tek başına çalışır; e-posta ile gönderilebilir, telefonda açılabilir."";" & vbLf
    s = s & "})();" & vbLf
    s = s & "</script>" & vbLf
    s = s & "</body>" & vbLf
    s = s & "</html>" & vbLf
    PS_05 = s
End Function

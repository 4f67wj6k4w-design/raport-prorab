// Финансовые итоги объекта для «Финального отчёта» панели руководства (папка f/).
// Запуск: EXEC_KEY=… node src/mkfin.js fin/kub16.json …   (исходные json в репозиторий НЕ класть)
// Имя файла = SHA-256("raport-fin-n|"+ключ панели+"|"+код)[0..24], ключ = SHA-256("raport-fin|"+ключ панели+"|"+код).
const fs=require("fs"),path=require("path"),{webcrypto:wc}=require("crypto");
const K=process.env.EXEC_KEY;if(!K||K.length<24){console.error("нужен EXEC_KEY");process.exit(1);}
const te=new TextEncoder(),hex=b=>Buffer.from(b).toString("hex");
(async()=>{for(const f of process.argv.slice(2)){
  const c=JSON.parse(fs.readFileSync(f,"utf8"));if(!c.pid)throw new Error(f+": нет pid");
  const nm=hex(await wc.subtle.digest("SHA-256",te.encode("raport-fin-n|"+K+"|"+c.pid))).slice(0,24);
  const key=await wc.subtle.importKey("raw",await wc.subtle.digest("SHA-256",te.encode("raport-fin|"+K+"|"+c.pid)),"AES-GCM",false,["encrypt"]);
  const iv=wc.getRandomValues(new Uint8Array(12));
  const d=new Uint8Array(await wc.subtle.encrypt({name:"AES-GCM",iv},key,te.encode(JSON.stringify(c))));
  const out=path.join(__dirname,"..","f",nm+".json");fs.mkdirSync(path.dirname(out),{recursive:true});
  fs.writeFileSync(out,JSON.stringify({e:1,iv:Buffer.from(iv).toString("base64"),d:Buffer.from(d).toString("base64")}));
  console.log("ok",c.pid,"→ f/"+nm+".json");
}})().catch(e=>{console.error(e.message);process.exit(1);});

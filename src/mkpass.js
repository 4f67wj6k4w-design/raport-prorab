// Зашифрованный паспорт объекта для папки o/: node src/mkpass.js objects/fil7.json
// Ключ шифрования = SHA-256("raport-cu|" + ключ базы + "|" + код объекта); без ссылки прораба файл не прочитать.
const fs=require("fs"),path=require("path"),{webcrypto:wc}=require("crypto");
(async()=>{for(const f of process.argv.slice(2)){
  const c=JSON.parse(fs.readFileSync(f,"utf8"));if(!c.cu||!c.db||!c.db.k)throw new Error(f+": нет cu или db.k");
  const pub=Object.assign({},c);delete pub.db;
  const raw=await wc.subtle.digest("SHA-256",new TextEncoder().encode("raport-cu|"+c.db.k+"|"+c.pid));
  const key=await wc.subtle.importKey("raw",raw,"AES-GCM",false,["encrypt"]);
  const iv=wc.getRandomValues(new Uint8Array(12));
  const d=new Uint8Array(await wc.subtle.encrypt({name:"AES-GCM",iv},key,new TextEncoder().encode(JSON.stringify(pub))));
  const out=path.join(__dirname,"..",c.cu);
  fs.writeFileSync(out,JSON.stringify({e:1,iv:Buffer.from(iv).toString("base64"),d:Buffer.from(d).toString("base64")}));
  console.log("ok",c.pid,"→",c.cu);
}})().catch(e=>{console.error(e.message);process.exit(1);});

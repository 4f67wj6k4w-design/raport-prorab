// Ключ прораба для объекта: FOREMAN_SECRET=... node src/objkey.js <pid> [gen]
// Должен совпадать с foremanKey_() в скрипте базы.
const c=require("crypto");
module.exports=(sec,pid,gen)=>c.createHmac("sha256",sec).update(String(pid)+"|"+(gen||1)).digest("base64url").slice(0,22);
if(require.main===module){const s=process.env.FOREMAN_SECRET;if(!s){console.error("нет FOREMAN_SECRET");process.exit(1);}console.log(module.exports(s,process.argv[2],+process.argv[3]||1));}

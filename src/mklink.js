// Ссылка прораба из JSON-настроек объекта: node src/mklink.js objects/fil7.json
const zlib=require("zlib"),fs=require("fs");const c=JSON.parse(fs.readFileSync(process.argv[2],"utf8"));
const z=zlib.deflateRawSync(Buffer.from(JSON.stringify(c)));
console.log("https://4f67wj6k4w-design.github.io/raport-prorab/#p=z"+z.toString("base64").replace(/\+/g,"-").replace(/\//g,"_").replace(/=+$/,""));

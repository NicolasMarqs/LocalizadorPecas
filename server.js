const http = require("http");
const fs = require("fs");
const path = require("path");

require("dotenv").config();

const { createClient } = require("@supabase/supabase-js");

const supabase = createClient(
    process.env.SUPABASE_URL,
    process.env.SUPABASE_SERVICE_KEY
);

async function testarSupabase() {
    const { data, error } = await supabase
        .from("pecas")
        .select("*")
        .limit(1);

    if (error) {
        console.log("Erro Supabase:", error.message);
    } else {
        console.log("Supabase conectado com sucesso!");
    }
}

testarSupabase();   

const formidable = require("formidable");
const XLSX = require("xlsx");


const PORTA = process.env.PORT || 3000;
const SENHA_ADMIN = process.env.SENHA_ADMIN;

const servidor = http.createServer(async (req, res) => {

if (req.method === "GET" && req.url.startsWith("/api/peca?codigo=")) {

    const codigo = decodeURIComponent(
        req.url.split("codigo=")[1] || ""
    ).trim();

    const { data, error } = await supabase
        .from("pecas")
        .select("codigo, descricao, local")
        .eq("codigo", codigo);

    res.writeHead(error ? 500 : 200, {
        "Content-Type": "application/json; charset=utf-8"
    });

    res.end(JSON.stringify({
        sucesso: !error,
        pecas: data || [],
        erro: error ? error.message : null
    }));

    return;
}

    if (req.method === "POST" && req.url === "/verificar-senha") {

    let corpo = "";

    req.on("data", parte => {
        corpo += parte;
    });

    req.on("end", () => {
        const dados = JSON.parse(corpo);

        if (dados.senha === SENHA_ADMIN) {
            res.writeHead(200, {
                "Content-Type": "application/json"
            });
            res.end(JSON.stringify({ correta: true }));
        } else {
            res.writeHead(401, {
                "Content-Type": "application/json"
            });
            res.end(JSON.stringify({ correta: false }));
        }
    });

    return;
}

       if (req.method === "POST" && req.url === "/atualizar") {

        const form = formidable.formidable({
            keepExtensions: true
        });

        form.parse(req, async (erro, campos, arquivos) => {
            
            if (erro) {
                res.writeHead(500);
                res.end("Erro ao receber a planilha.");
                return;
            }

            const senhaRecebida = Array.isArray(campos.senha)
    ? campos.senha[0]
    : campos.senha;

if (senhaRecebida !== SENHA_ADMIN) {
    res.writeHead(403, {
        "Content-Type": "text/plain; charset=utf-8"
    });
    res.end("Senha incorreta.");
    return;
}

            console.log("Planilha recebida pelo servidor!");

            const arquivoExcel = Array.isArray(arquivos.planilha)
    ? arquivos.planilha[0]
    : arquivos.planilha;

if (!arquivoExcel) {
    res.writeHead(400);
    res.end("Nenhuma planilha foi enviada.");
    return;
}

const workbook = XLSX.readFile(arquivoExcel.filepath);

const primeiraPlanilha =
    workbook.Sheets[workbook.SheetNames[0]];

const dadosPlanilha =
    XLSX.utils.sheet_to_json(primeiraPlanilha);

console.log("Peças encontradas:", dadosPlanilha.length);

const linhasCSV = [
    "codigo,descricao,local",
    ...dadosPlanilha.map(item => {
const codigo = String(item["Código do Item"] ?? "").trim();
const descricao = String(item["Descrição"] ?? "").trim();
const local = String(item["Locação"] ?? "").trim(); 

        return `${codigo},${descricao},${local}`;
    })
];

const pecasParaBanco = dadosPlanilha
    .map(item => ({
        codigo: String(item["Código do Item"] ?? "").trim(),
        descricao: String(item["Descrição"] ?? "").trim(),
        local: String(item["Locação"] ?? "").trim()
    }))
    .filter(item => item.codigo !== "");

const pecasSemDuplicados = [
    ...new Map(
        pecasParaBanco.map(item => [item.codigo, item])
    ).values()
];

console.log(
    "Peças sem códigos duplicados:",
    pecasSemDuplicados.length
);


for (let i = 0; i < pecasSemDuplicados.length; i += 1000) {
    const lote = pecasSemDuplicados.slice(i, i + 1000);

    const { error: erroInserir } = await supabase
        .from("pecas")
        .upsert(lote, { onConflict: "codigo" });

    if (erroInserir) {
        console.log("Erro ao inserir no Supabase:", erroInserir.message);

        res.writeHead(500, {
            "Content-Type": "text/plain; charset=utf-8"
        });

        res.end("Erro ao atualizar banco de peças.");
        return;
    }
}

console.log("Supabase atualizado:", pecasSemDuplicados.length, "peças");

const conteudoCSV = linhasCSV.join("\n");

fs.writeFileSync(
    path.join(__dirname, "pecas.csv"),
    conteudoCSV,
    "utf8"
);

console.log("pecas.csv atualizado!");

            res.writeHead(200, {
                "Content-Type": "text/plain; charset=utf-8"
            });

            res.end("Planilha recebida!");
        });

        return;
    }

    let arquivo = req.url === "/" ? "index.html" : req.url.substring(1);

    const caminhoArquivo = path.join(__dirname, arquivo);

    fs.readFile(caminhoArquivo, (erro, conteudo) => {

        if (erro) {
            res.writeHead(404);
            res.end("Arquivo não encontrado");
            return;
        }

        res.writeHead(200);
        res.end(conteudo);
    });

});

servidor.listen(PORTA, "0.0.0.0", () => {
    console.log("Servidor do Localizador de Peças iniciado!");
    console.log("Porta: " + PORTA);
});
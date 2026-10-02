const http = require("http");
const fs = require("fs");
const path = require("path");

require("dotenv").config();

const { createClient } = require("@supabase/supabase-js");
const formidable = require("formidable");
const XLSX = require("xlsx");

const supabase = createClient(
    process.env.SUPABASE_URL,
    process.env.SUPABASE_SERVICE_KEY
);

const PORTA = process.env.PORT || 3000;
const SENHA_ADMIN = process.env.SENHA_ADMIN;


// =====================================================
// TESTAR CONEXÃO COM SUPABASE
// =====================================================

async function testarSupabase() {

    const { error } = await supabase
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


// =====================================================
// SERVIDOR
// =====================================================

const servidor = http.createServer(async (req, res) => {


    // =================================================
    // CONSULTAR PEÇA
    // =================================================

    if (
        req.method === "GET" &&
        req.url.startsWith("/api/peca?codigo=")
    ) {

        const codigo = decodeURIComponent(
            req.url.split("codigo=")[1] || ""
        ).trim();

        const { data, error } = await supabase
            .from("pecas")
            .select("codigo, descricao, local")
            .eq("codigo", codigo);

        res.writeHead(
            error ? 500 : 200,
            {
                "Content-Type":
                    "application/json; charset=utf-8"
            }
        );

        res.end(
            JSON.stringify({
                sucesso: !error,
                pecas: data || [],
                erro: error
                    ? error.message
                    : null
            })
        );

        return;
    }


    // =================================================
    // VERIFICAR SENHA DO ADMINISTRADOR
    // =================================================

    if (
        req.method === "POST" &&
        req.url === "/verificar-senha"
    ) {

        let corpo = "";

        req.on("data", parte => {
            corpo += parte;
        });

        req.on("end", () => {

            try {

                const dados =
                    JSON.parse(corpo);

                if (
                    dados.senha ===
                    SENHA_ADMIN
                ) {

                    res.writeHead(
                        200,
                        {
                            "Content-Type":
                                "application/json"
                        }
                    );

                    res.end(
                        JSON.stringify({
                            correta: true
                        })
                    );

                } else {

                    res.writeHead(
                        401,
                        {
                            "Content-Type":
                                "application/json"
                        }
                    );

                    res.end(
                        JSON.stringify({
                            correta: false
                        })
                    );
                }

            } catch (erro) {

                res.writeHead(
                    400,
                    {
                        "Content-Type":
                            "application/json"
                    }
                );

                res.end(
                    JSON.stringify({
                        correta: false
                    })
                );
            }
        });

        return;
    }


    // =================================================
    // ATUALIZAR BASE DE PEÇAS
    // =================================================

    if (
        req.method === "POST" &&
        req.url === "/atualizar"
    ) {

        const form =
            formidable.formidable({
                keepExtensions: true
            });

        form.parse(
            req,
            async (
                erro,
                campos,
                arquivos
            ) => {

                if (erro) {

                    console.log(
                        "Erro no upload:",
                        erro
                    );

                    res.writeHead(
                        500,
                        {
                            "Content-Type":
                                "text/plain; charset=utf-8"
                        }
                    );

                    res.end(
                        "Erro ao receber a planilha."
                    );

                    return;
                }


                // =====================================
                // VALIDAR SENHA
                // =====================================

                const senhaRecebida =
                    Array.isArray(
                        campos.senha
                    )
                        ? campos.senha[0]
                        : campos.senha;

                if (
                    senhaRecebida !==
                    SENHA_ADMIN
                ) {

                    res.writeHead(
                        403,
                        {
                            "Content-Type":
                                "text/plain; charset=utf-8"
                        }
                    );

                    res.end(
                        "Senha incorreta."
                    );

                    return;
                }


                // =====================================
                // LOCALIZAR ARQUIVO
                // =====================================

                const arquivoExcel =
                    Array.isArray(
                        arquivos.planilha
                    )
                        ? arquivos.planilha[0]
                        : arquivos.planilha;

                if (!arquivoExcel) {

                    res.writeHead(
                        400,
                        {
                            "Content-Type":
                                "text/plain; charset=utf-8"
                        }
                    );

                    res.end(
                        "Nenhuma planilha foi enviada."
                    );

                    return;
                }


                try {

                    console.log(
                        "Planilha recebida pelo servidor!"
                    );


                    // =================================
                    // LER EXCEL
                    // =================================

                    const workbook =
                        XLSX.readFile(
                            arquivoExcel.filepath
                        );

                    const primeiraPlanilha =
                        workbook.Sheets[
                            workbook.SheetNames[0]
                        ];

                    const dadosPlanilha =
                        XLSX.utils.sheet_to_json(
                            primeiraPlanilha
                        );

                    console.log(
                        "Linhas encontradas:",
                        dadosPlanilha.length
                    );


                    // =================================
                    // CONVERTER PARA FORMATO DO BANCO
                    // =================================

                    const pecasParaBanco =
                        dadosPlanilha
                            .map(item => ({

                                codigo: String(
                                    item["Cód.Item"] ??
                                    ""
                                ).trim(),

                                descricao: String(
                                    item["Descrição"] ??
                                    ""
                                ).trim(),

                                local: String(
                                    item["Locação"] ??
                                    ""
                                ).trim()

                            }))
                            .filter(
                                item =>
                                    item.codigo !== ""
                            );


                    // =================================
                    // REMOVER CÓDIGOS DUPLICADOS
                    // =================================

                    const pecasSemDuplicados =
                        [
                            ...new Map(
                                pecasParaBanco.map(
                                    item => [
                                        item.codigo,
                                        item
                                    ]
                                )
                            ).values()
                        ];


                    console.log(
                        "Peças válidas:",
                        pecasSemDuplicados.length
                    );


                    // =================================
                    // PROTEÇÃO CONTRA PLANILHA INVÁLIDA
                    // =================================

                    if (
                        pecasSemDuplicados.length === 0
                    ) {

                        res.writeHead(
                            400,
                            {
                                "Content-Type":
                                    "text/plain; charset=utf-8"
                            }
                        );

                        res.end(
                            "A planilha não contém peças válidas. " +
                            "A base atual foi mantida."
                        );

                        return;
                    }


                    // =================================
                    // APAGAR BASE ANTIGA
                    // =================================

                    console.log(
                        "Limpando base antiga..."
                    );

                    const {
                        error: erroApagar
                    } = await supabase
                        .from("pecas")
                        .delete()
                        .not(
                            "id",
                            "is",
                            null
                        );


                    if (erroApagar) {

                        console.log(
                            "Erro ao limpar Supabase:",
                            erroApagar.message
                        );

                        res.writeHead(
                            500,
                            {
                                "Content-Type":
                                    "text/plain; charset=utf-8"
                            }
                        );

                        res.end(
                            "Erro ao limpar a base antiga."
                        );

                        return;
                    }


                    console.log(
                        "Base antiga removida."
                    );


                    // =================================
                    // INSERIR NOVA BASE
                    // =================================

                    let totalInserido = 0;


                    for (
                        let i = 0;
                        i <
                        pecasSemDuplicados.length;
                        i += 500
                    ) {

                        const lote =
                            pecasSemDuplicados.slice(
                                i,
                                i + 500
                            );


                        const {
                            error: erroInserir
                        } = await supabase
                            .from("pecas")
                            .insert(lote);


                        if (erroInserir) {

                            console.log(
                                "Erro ao inserir no Supabase:",
                                erroInserir.message
                            );

                            res.writeHead(
                                500,
                                {
                                    "Content-Type":
                                        "text/plain; charset=utf-8"
                                }
                            );

                            res.end(
                                "Erro ao gravar a nova base " +
                                "após " +
                                totalInserido +
                                " peças."
                            );

                            return;
                        }


                        totalInserido +=
                            lote.length;


                        console.log(
                            "Inseridas:",
                            totalInserido,
                            "de",
                            pecasSemDuplicados.length
                        );
                    }


                    // =================================
                    // CONFERIR QUANTIDADE
                    // =================================

                    const {
                        count,
                        error: erroContagem
                    } = await supabase
                        .from("pecas")
                        .select(
                            "*",
                            {
                                count: "exact",
                                head: true
                            }
                        );


                    if (erroContagem) {

                        console.log(
                            "Erro na conferência:",
                            erroContagem.message
                        );

                        res.writeHead(
                            500,
                            {
                                "Content-Type":
                                    "text/plain; charset=utf-8"
                            }
                        );

                        res.end(
                            "A base foi enviada, mas não foi possível " +
                            "conferir a quantidade."
                        );

                        return;
                    }


                    if (
                        count !==
                        pecasSemDuplicados.length
                    ) {

                        console.log(
                            "Quantidade diferente.",
                            "Esperado:",
                            pecasSemDuplicados.length,
                            "Banco:",
                            count
                        );

                        res.writeHead(
                            500,
                            {
                                "Content-Type":
                                    "text/plain; charset=utf-8"
                            }
                        );

                        res.end(
                            "A quantidade gravada no banco " +
                            "não corresponde à planilha."
                        );

                        return;
                    }


                    // =================================
                    // SUCESSO
                    // =================================

                    console.log(
                        "Supabase atualizado com sucesso:",
                        count,
                        "peças"
                    );


                    res.writeHead(
                        200,
                        {
                            "Content-Type":
                                "text/plain; charset=utf-8"
                        }
                    );


                    res.end(
                        `Base atualizada com sucesso: ${count} peças.`
                    );


                } catch (erro) {

                    console.log(
                        "Erro durante atualização:",
                        erro
                    );

                    res.writeHead(
                        500,
                        {
                            "Content-Type":
                                "text/plain; charset=utf-8"
                        }
                    );

                    res.end(
                        "Erro durante a atualização da base."
                    );
                }
            }
        );

        return;
    }


    // =================================================
    // SERVIR ARQUIVOS DO SITE
    // =================================================

    let arquivo =
        req.url === "/"
            ? "index.html"
            : req.url.substring(1);


    const caminhoArquivo =
        path.join(
            __dirname,
            arquivo
        );


    fs.readFile(
        caminhoArquivo,
        (
            erro,
            conteudo
        ) => {

            if (erro) {

                res.writeHead(404);

                res.end(
                    "Arquivo não encontrado"
                );

                return;
            }


            res.writeHead(200);

            res.end(conteudo);
        }
    );
});


// =====================================================
// INICIAR SERVIDOR
// =====================================================

servidor.listen(
    PORTA,
    "0.0.0.0",
    () => {

        console.log(
            "Servidor do Localizador de Peças iniciado!"
        );

        console.log(
            "Porta: " + PORTA
        );
    }
);

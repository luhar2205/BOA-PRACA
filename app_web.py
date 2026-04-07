import os
import shutil
import uuid
import zipfile
import json
from datetime import datetime
from fastapi import FastAPI, File, UploadFile, BackgroundTasks
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
import uvicorn

# Importando o motor do Harley
from core.processor import rodar_automacao_total
from core.exporter import salvar_excel_kildere

app = FastAPI(title="Boa Praça OS - Enterprise")

# =====================================================================
# 🛡️ ARQUITETURA E DIRETÓRIOS
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMP_DIR = os.path.join(BASE_DIR, "temp_files")
os.makedirs(TEMP_DIR, exist_ok=True)
CAMINHO_BASE_SANKHYA = os.path.join(BASE_DIR, "PRODUTOS_KILDERE.xlsx")

def limpar_pasta_temporaria(caminho_pasta: str):
    try:
        if os.path.exists(caminho_pasta):
            shutil.rmtree(caminho_pasta)
    except Exception: pass

# =====================================================================
# 🎨 FRONT-END AVANÇADO (Single Page Application com Vue/JS Vanilla)
# =====================================================================
HTML_CONTENT = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Boa Praça OS | Hub Operacional</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;800&display=swap');
        body { font-family: 'Inter', sans-serif; background-color: #f3f4f6; }
        .drag-active { border-color: #3b82f6 !important; background-color: #eff6ff !important; }
        .glass-panel { background: rgba(255, 255, 255, 0.95); backdrop-filter: blur(10px); }
        
        /* Custom Scrollbar para a tabela */
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: #f1f1f1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: #c1c1c1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #a8a8a8; }
    </style>
</head>
<body class="h-screen flex overflow-hidden text-gray-800">

    <aside class="w-64 bg-[#1F4068] text-white flex flex-col shadow-2xl z-20 relative">
        <div class="p-6 border-b border-blue-800/50">
            <h1 class="text-2xl font-extrabold tracking-wider flex items-center">
                <i class="fa-solid fa-anchor mr-3 text-blue-400"></i> BOA PRAÇA
            </h1>
            <p class="text-xs text-blue-300 mt-1 uppercase tracking-widest font-semibold">Operating System</p>
        </div>
        
        <nav class="flex-1 p-4 space-y-2">
            <a href="#" class="flex items-center p-3 bg-blue-800/50 text-white rounded-lg transition font-semibold">
                <i class="fa-solid fa-microchip w-6 text-blue-400"></i> Motor Sankhya UA
            </a>
            <a href="#" onclick="alert('Módulo em desenvolvimento! Aqui entrarão as abas do Resumão.')" class="flex items-center p-3 text-gray-400 hover:bg-blue-800/30 hover:text-white rounded-lg transition font-semibold">
                <i class="fa-solid fa-book-atlas w-6"></i> Wiki Operacional
            </a>
            <a href="#" class="flex items-center p-3 text-gray-400 hover:bg-blue-800/30 hover:text-white rounded-lg transition font-semibold">
                <i class="fa-solid fa-chart-pie w-6"></i> Analytics & Relatórios
            </a>
        </nav>
        
        <div class="p-4 border-t border-blue-800/50">
            <div class="flex items-center">
                <div class="w-8 h-8 rounded-full bg-blue-500 flex items-center justify-center text-sm font-bold shadow-lg">HL</div>
                <div class="ml-3">
                    <p class="text-sm font-bold">Harley Luiz</p>
                    <p class="text-xs text-blue-300">Administrador</p>
                </div>
            </div>
        </div>
    </aside>

    <main class="flex-1 flex flex-col h-screen overflow-hidden bg-gray-50 relative">
        
        <header class="h-16 bg-white shadow-sm flex items-center justify-between px-8 z-10">
            <h2 class="text-xl font-bold text-gray-700">Processamento Inteligente de Notas Fiscais</h2>
            <div class="flex space-x-4">
                <button class="text-gray-400 hover:text-blue-600 transition"><i class="fa-solid fa-bell"></i></button>
                <button class="text-gray-400 hover:text-blue-600 transition"><i class="fa-solid fa-gear"></i></button>
            </div>
        </header>

        <div class="flex-1 overflow-y-auto p-8">
            
            <div id="uploadScreen" class="max-w-4xl mx-auto mt-4">
                <div class="glass-panel rounded-2xl shadow-xl p-10 border border-gray-100">
                    <div class="text-center mb-8">
                        <h3 class="text-2xl font-extrabold text-[#1F4068]">Importar Lote (XML/ZIP)</h3>
                        <p class="text-gray-500 mt-2">O motor fará a varredura e o cruzamento com a base Sankhya automaticamente.</p>
                    </div>
                    
                    <form id="uploadForm">
                        <div id="dropzone" class="border-2 border-dashed border-gray-300 rounded-xl p-16 transition duration-300 relative cursor-pointer hover:bg-blue-50 flex flex-col items-center justify-center bg-gray-50/50">
                            <input type="file" id="fileInput" accept=".xml, .zip" multiple required class="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-50"/>
                            <div class="w-20 h-20 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center mb-4 shadow-sm">
                                <i class="fa-solid fa-cloud-arrow-up text-3xl"></i>
                            </div>
                            <h3 class="font-bold text-lg text-gray-700 mb-1">Arraste os arquivos para cá</h3>
                            <p class="text-sm text-gray-400">Ou clique para selecionar na sua máquina</p>
                            
                            <div id="file-list" class="hidden mt-6 w-full max-w-md bg-white p-4 rounded-lg shadow-sm border border-gray-100 text-left relative z-40">
                                <p class="text-xs font-bold text-blue-600 uppercase tracking-wider mb-2">Arquivos na fila:</p>
                                <p id="file-names" class="text-sm text-gray-600 truncate"></p>
                            </div>
                        </div>
                        
                        <button type="submit" id="submitBtn" class="w-full bg-[#1F4068] hover:bg-[#162d4a] text-white font-bold py-4 rounded-xl shadow-lg transition duration-300 flex justify-center items-center text-lg mt-6">
                            <i class="fa-solid fa-bolt text-yellow-400 mr-3"></i> INICIAR CRUZAMENTO DE DADOS
                        </button>
                    </form>
                </div>
            </div>

            <div id="loadingScreen" class="hidden h-full flex flex-col items-center justify-center">
                <div class="relative">
                    <div class="w-24 h-24 border-8 border-gray-200 border-t-[#1F4068] rounded-full animate-spin"></div>
                    <i class="fa-solid fa-microchip absolute top-1/2 left-1/2 transform -translate-x-1/2 -translate-y-1/2 text-2xl text-[#1F4068]"></i>
                </div>
                <h2 class="text-2xl font-bold text-[#1F4068] mt-8 mb-2">Processando Lote...</h2>
                <p class="text-gray-500">Extraindo dados, limpando EANs e calculando conversões de UA.</p>
            </div>

            <div id="resultScreen" class="hidden max-w-7xl mx-auto space-y-6">
                
                <div class="flex justify-between items-end">
                    <div>
                        <h2 class="text-3xl font-extrabold text-[#1F4068]">Relatório de Operação</h2>
                        <p class="text-gray-500 mt-1">Análise concluída com sucesso.</p>
                    </div>
                    <div class="flex space-x-3">
                        <button onclick="window.location.reload()" class="px-6 py-2.5 bg-white border border-gray-300 text-gray-700 font-bold rounded-lg hover:bg-gray-50 shadow-sm transition">
                            <i class="fa-solid fa-rotate-left mr-2"></i> Novo Lote
                        </button>
                        <a id="downloadBtn" href="#" class="px-6 py-2.5 bg-green-600 hover:bg-green-700 text-white font-bold rounded-lg shadow-lg transition flex items-center">
                            <i class="fa-solid fa-file-excel mr-2"></i> Exportar Excel
                        </a>
                    </div>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                    <div class="bg-white rounded-xl shadow-sm p-6 border border-gray-100 flex items-center">
                        <div class="w-14 h-14 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center text-2xl mr-4"><i class="fa-solid fa-boxes-stacked"></i></div>
                        <div>
                            <p class="text-sm font-bold text-gray-400 uppercase">Total de Itens</p>
                            <p id="insight-total" class="text-3xl font-black text-gray-800">0</p>
                        </div>
                    </div>
                    <div class="bg-white rounded-xl shadow-sm p-6 border border-green-100 flex items-center relative overflow-hidden">
                        <div class="absolute right-0 top-0 w-2 h-full bg-green-500"></div>
                        <div class="w-14 h-14 rounded-full bg-green-100 text-green-600 flex items-center justify-center text-2xl mr-4"><i class="fa-solid fa-check-double"></i></div>
                        <div>
                            <p class="text-sm font-bold text-gray-400 uppercase">Match Perfeito</p>
                            <p id="insight-match" class="text-3xl font-black text-gray-800">0</p>
                        </div>
                    </div>
                    <div class="bg-white rounded-xl shadow-sm p-6 border border-red-100 flex items-center relative overflow-hidden">
                        <div class="absolute right-0 top-0 w-2 h-full bg-red-500"></div>
                        <div class="w-14 h-14 rounded-full bg-red-100 text-red-600 flex items-center justify-center text-2xl mr-4"><i class="fa-solid fa-triangle-exclamation"></i></div>
                        <div>
                            <p class="text-sm font-bold text-gray-400 uppercase">Revisão Necessária</p>
                            <p id="insight-error" class="text-3xl font-black text-gray-800">0</p>
                        </div>
                    </div>
                </div>

                <div class="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden flex flex-col">
                    <div class="p-4 border-b border-gray-200 bg-gray-50 flex justify-between items-center">
                        <h3 class="font-bold text-gray-700"><i class="fa-solid fa-table-list mr-2"></i> Detalhamento por Item</h3>
                        <div class="relative">
                            <i class="fa-solid fa-magnifying-glass absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400"></i>
                            <input type="text" id="searchInput" onkeyup="filtrarTabela()" placeholder="Buscar produto ou código..." class="pl-10 pr-4 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 w-64">
                        </div>
                    </div>
                    
                    <div class="overflow-x-auto max-h-[500px]">
                        <table class="w-full text-left border-collapse" id="dataTable">
                            <thead class="bg-gray-100 text-gray-600 text-xs uppercase tracking-wider sticky top-0 shadow-sm">
                                <tr>
                                    <th class="p-4 font-bold">Cód Sankhya</th>
                                    <th class="p-4 font-bold">Descrição (Matriz)</th>
                                    <th class="p-4 font-bold text-right">Qtd NF</th>
                                    <th class="p-4 font-bold text-right">Peso (uTrib)</th>
                                    <th class="p-4 font-bold text-center bg-blue-50 text-blue-800 border-x border-blue-100">Cálculo UA</th>
                                    <th class="p-4 font-bold text-center">Status de Inteligência</th>
                                </tr>
                            </thead>
                            <tbody id="tableBody" class="text-sm divide-y divide-gray-100">
                                </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </main>

    <script>
        const fileInput = document.getElementById('fileInput');
        const dropzone = document.getElementById('dropzone');
        const fileList = document.getElementById('file-list');
        const fileNames = document.getElementById('file-names');
        
        const uploadScreen = document.getElementById('uploadScreen');
        const loadingScreen = document.getElementById('loadingScreen');
        const resultScreen = document.getElementById('resultScreen');

        fileInput.addEventListener('dragenter', () => dropzone.classList.add('drag-active'));
        fileInput.addEventListener('dragleave', () => dropzone.classList.remove('drag-active'));
        fileInput.addEventListener('drop', () => dropzone.classList.remove('drag-active'));

        fileInput.addEventListener('change', () => {
            if(fileInput.files.length > 0) {
                fileList.classList.remove('hidden');
                fileNames.textContent = Array.from(fileInput.files).map(f => f.name).join(', ');
            }
        });

        function filtrarTabela() {
            let input = document.getElementById("searchInput").value.toUpperCase();
            let trs = document.getElementById("dataTable").getElementsByTagName("tr");
            for (let i = 1; i < trs.length; i++) {
                let text = trs[i].textContent || trs[i].innerText;
                trs[i].style.display = text.toUpperCase().indexOf(input) > -1 ? "" : "none";
            }
        }

        document.getElementById('uploadForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            uploadScreen.classList.add('hidden');
            loadingScreen.classList.remove('hidden');

            const formData = new FormData();
            for (let i = 0; i < fileInput.files.length; i++) {
                formData.append('arquivos', fileInput.files[i]);
            }

            try {
                const response = await fetch('/api/processar', {
                    method: 'POST',
                    body: formData
                });
                
                const result = await response.json();
                
                if(!response.ok) {
                    alert(result.error || "Ocorreu um erro no servidor.");
                    window.location.reload();
                    return;
                }

                document.getElementById('insight-total').textContent = result.estatisticas.total;
                document.getElementById('insight-match').textContent = result.estatisticas.sucessos;
                document.getElementById('insight-error').textContent = result.estatisticas.erros;
                document.getElementById('downloadBtn').href = `/api/download/${result.session_id}`;

                const tbody = document.getElementById('tableBody');
                tbody.innerHTML = '';
                
                result.dados.forEach(item => {
                    const row = document.createElement('tr');
                    row.className = "hover:bg-gray-50 transition";
                    
                    let statusBadge = '';
                    if(item.MATCH.includes('✅')) {
                        statusBadge = `<span class="px-3 py-1 bg-green-100 text-green-700 font-bold rounded-full text-xs">${item.MATCH}</span>`;
                    } else if(item.MATCH.includes('🔶')) {
                        statusBadge = `<span class="px-3 py-1 bg-yellow-100 text-yellow-700 font-bold rounded-full text-xs">${item.MATCH}</span>`;
                    } else {
                        statusBadge = `<span class="px-3 py-1 bg-red-100 text-red-700 font-bold rounded-full text-xs"><i class="fa-solid fa-triangle-exclamation mr-1"></i> ${item.MATCH}</span>`;
                    }

                    row.innerHTML = `
                        <td class="p-4 font-bold text-gray-700">${item['CÓD SANKHYA'] || '-'}</td>
                        <td class="p-4">
                            <p class="font-bold text-[#1F4068]">${item['DESC SANKHYA'] || 'Produto não encontrado'}</p>
                            <p class="text-xs text-gray-400 mt-1">NF: ${item['PRODUTO XML']}</p>
                        </td>
                        <td class="p-4 text-right text-gray-600">${item['QTD NOTA']} ${item['UN NOTA']}</td>
                        <td class="p-4 text-right font-semibold text-gray-800">${item['Peso NF (uTrib)']}</td>
                        <td class="p-4 text-center bg-blue-50/50 border-x border-blue-50 font-mono font-bold text-blue-700">${item['UA / CONVERSÃO']}</td>
                        <td class="p-4 text-center">${statusBadge}</td>
                    `;
                    tbody.appendChild(row);
                });

                loadingScreen.classList.add('hidden');
                resultScreen.classList.remove('hidden');

            } catch (error) {
                alert("Erro ao conectar com o motor de inteligência.");
                window.location.reload();
            }
        });
    </script>
</body>
</html>
"""

# =====================================================================
# ⚙️ API REST (Back-end)
# =====================================================================

@app.get("/", response_class=HTMLResponse)
async def home():
    return HTMLResponse(content=HTML_CONTENT)

@app.post("/api/processar")
async def api_processar(
    background_tasks: BackgroundTasks,
    arquivos: list[UploadFile] = File(...)
):
    if not os.path.exists(CAMINHO_BASE_SANKHYA):
        return JSONResponse(status_code=500, content={"error": "Base PRODUTOS_KILDERE.xlsx não encontrada no servidor."})

    sessao_id = uuid.uuid4().hex
    pasta_sessao = os.path.join(TEMP_DIR, sessao_id)
    pasta_xmls = os.path.join(pasta_sessao, "xmls")
    os.makedirs(pasta_xmls, exist_ok=True)
        
    for arquivo in arquivos:
        extensao = os.path.splitext(arquivo.filename)[1].lower()
        caminho_temp = os.path.join(pasta_sessao, arquivo.filename)
        
        with open(caminho_temp, "wb") as buffer:
            shutil.copyfileobj(arquivo.file, buffer)
            
        if extensao == '.zip':
            try:
                with zipfile.ZipFile(caminho_temp, 'r') as zip_ref:
                    for file_info in zip_ref.infolist():
                        if file_info.filename.lower().endswith('.xml'):
                            zip_ref.extract(file_info, pasta_xmls)
            except Exception: pass
        elif extensao == '.xml':
            shutil.move(caminho_temp, os.path.join(pasta_xmls, arquivo.filename))
            
    df_resultado = rodar_automacao_total(pasta_xmls, CAMINHO_BASE_SANKHYA)
    
    nome_arquivo = f"Conversao_Sankhya_{datetime.now().strftime('%d_%m_%Hh%M')}.xlsx"
    caminho_excel = os.path.join(pasta_sessao, nome_arquivo)
    salvar_excel_kildere(df_resultado, caminho_excel)
    
    registros_json = df_resultado.to_dict(orient="records")
    total_itens = len(registros_json)
    sucessos = sum(1 for r in registros_json if "✅" in str(r.get("MATCH", "")))
    inteligentes = sum(1 for r in registros_json if "🔶" in str(r.get("MATCH", "")))
    erros = sum(1 for r in registros_json if "❌" in str(r.get("MATCH", "")))
    
    return JSONResponse(content={
        "session_id": f"{sessao_id}/{nome_arquivo}",
        "estatisticas": {
            "total": total_itens,
            "sucessos": sucessos + inteligentes,
            "erros": erros
        },
        "dados": registros_json
    })

@app.get("/api/download/{sessao_id}/{nome_arquivo}")
async def api_download(sessao_id: str, nome_arquivo: str, background_tasks: BackgroundTasks):
    pasta_sessao = os.path.join(TEMP_DIR, sessao_id)
    caminho_arquivo = os.path.join(pasta_sessao, nome_arquivo)
    
    background_tasks.add_task(limpar_pasta_temporaria, pasta_sessao)
    
    return FileResponse(
        path=caminho_arquivo, 
        filename=nome_arquivo,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    # 🌩️ O Railway preenche a variável de ambiente PORT. Se rodar no PC, usa 8000.
    porta = int(os.environ.get("PORT", 8000))
    print(f"🚀 BOA PRAÇA OS - PREPARADO PARA NUVEM (PORTA {porta})")
    
    # host="0.0.0.0" permite conexões externas do servidor do Railway
    uvicorn.run(app, host="0.0.0.0", port=porta)
let dadosPlanilha = [];
let dadosAgrupados = [];
let dadosOriginais = [];
let arquivoSelecionado = null;

// Captura o arquivo escolhido
document.getElementById("inputExcel").addEventListener("change", (e) => {
  arquivoSelecionado = e.target.files[0];
});

// Botão para carregar planilha
function carregarPlanilha() {
  if (!arquivoSelecionado) {
    alert("Selecione um arquivo primeiro!");
    return;
  }

  const reader = new FileReader();
  reader.onload = (event) => {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    let dados = XLSX.utils.sheet_to_json(sheet);

    // Normaliza nomes das colunas
    dadosPlanilha = dados.map(linha => {
      return {
        Pedido: linha["Pedido"] || linha["pedido"] || "",
        Descricao: linha["Descrição"] || linha["Descricao"] || linha["Produto"] || "",
        QtdSolicitada: Number(linha["Qtd.Solicitada"] || linha["QtdSolicitada"] || linha["Quantidade"] || 0),
        DataEntrega: tratarData(linha["Data de entrega"] || linha["DataEntrega"] || linha["Entrega"] || "")
      };
    });

    mostrarTabela("tabelaPedidos", dadosPlanilha);
  };
  reader.readAsArrayBuffer(arquivoSelecionado);
}

// Conversão de data do Excel
function tratarData(valor) {
  if (!valor) return "";
  if (typeof valor === "number") {
    const dataBase = new Date(1899, 11, 30);
    const data = new Date(dataBase.getTime() + valor * 24 * 60 * 60 * 1000);
    return data.toLocaleDateString("pt-BR");
  }
  return valor.toString();
}

// Mostrar tabela simples (Pedidos)
function mostrarTabela(id, dados) {
  if (!dados.length) {
    document.getElementById(id).innerHTML = "<p style='color:gray'>Nenhum dado disponível</p>";
    return;
  }

  let html = "<table><tr>";
  Object.keys(dados[0]).forEach(col => {
    html += `<th>${col}</th>`;
  });
  html += "</tr>";

  dados.forEach(row => {
    html += "<tr>";
    Object.keys(row).forEach(col => {
      html += `<td>${row[col]}</td>`;
    });
    html += "</tr>";
  });

  html += "</table>";
  document.getElementById(id).innerHTML = html;
}

// Calcular agrupamento com pedidos limitados na tela
function calcular() {
  if (!dadosPlanilha.length) {
    alert("Carregue a planilha primeiro!");
    return;
  }

  let agrupado = {};

  dadosPlanilha.forEach(item => {
    const chave = item.Descricao;
    if (!agrupado[chave]) {
      agrupado[chave] = {
        Pedidos: [],
        Descrição: item.Descricao,
        DataEntrega: null,
        QtdTotal: 0
      };
    }

    // adiciona pedidos
    agrupado[chave].Pedidos.push(item.Pedido);

    // data mais próxima
    if (item.DataEntrega) {
      let dataAtual = new Date(item.DataEntrega.split("/").reverse().join("-"));
      if (!agrupado[chave].DataEntrega) {
        agrupado[chave].DataEntrega = item.DataEntrega;
      } else {
        let dataExistente = new Date(agrupado[chave].DataEntrega.split("/").reverse().join("-"));
        if (dataAtual < dataExistente) agrupado[chave].DataEntrega = item.DataEntrega;
      }
    }

    // soma quantidade
    agrupado[chave].QtdTotal += item.QtdSolicitada;
  });

  // transforma para exibição + exportação
  dadosAgrupados = Object.values(agrupado).map(item => {
    let pedidosTela = item.Pedidos.length > 3 ? item.Pedidos.slice(0,3).join("/") + "..." : item.Pedidos.join("/");
    return {
      Pedidos: pedidosTela,
      PedidosExport: item.Pedidos.join("/"),
      Descrição: item.Descrição,
      "Data de entrega": item.DataEntrega,
      "Qtd.Total": item.QtdTotal.toLocaleString("pt-BR", { minimumFractionDigits: 2 })
    };
  });

  dadosOriginais = [...dadosAgrupados]; // Salva os dados originais para filtro
  mostrarTabelaTela("tabelaCalcular", dadosAgrupados);
}

// Mostrar tabela de cálculo na tela (limita pedidos)
function mostrarTabelaTela(id, dados) {
  if (!dados.length) {
    document.getElementById(id).innerHTML = "<p style='color:gray'>Nenhum dado disponível</p>";
    return;
  }

  let html = "<table><tr>";
  Object.keys(dados[0]).forEach(col => {
    if(col !== "PedidosExport") html += `<th>${col}</th>`;
  });
  html += "</tr>";

  dados.forEach(row => {
    html += "<tr>";
    Object.keys(row).forEach(col => {
      if(col !== "PedidosExport") {
        if(col === "Pedidos") html += `<td class="Pedidos">${row[col]}</td>`;
        else html += `<td>${row[col]}</td>`;
      }
    });
    html += "</tr>";
  });

  html += "</table>";
  document.getElementById(id).innerHTML = html;
}

// Exportar Excel com todos os pedidos
function exportarExcel() {
  if (!dadosAgrupados.length) {
    alert("Nenhum cálculo realizado ainda!");
    return;
  }

  const exportData = dadosAgrupados.map(item => ({
    Pedidos: item.PedidosExport,
    Descrição: item.Descrição,
    "Data de entrega": item["Data de entrega"],
    "Qtd.Total": item["Qtd.Total"]
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(exportData);
  XLSX.utils.book_append_sheet(wb, ws, "Calculo");
  XLSX.writeFile(wb, "resultado_mrp.xlsx");
}

// Filtrar pedidos por número do pedido
function filtrarPedidos() {
  if (!dadosOriginais.length) {
    return; // Nenhum dado para filtrar
  }

  const filtro = document.getElementById("filtroPedidos").value.toLowerCase();
  const filtrados = dadosOriginais.filter(item =>
    item.Pedidos.toLowerCase().includes(filtro)
  );

  mostrarTabelaTela("tabelaCalcular", filtrados);
}

// Alternar abas
function abrirAba(evt, nomeAba) {
  const conteudos = document.getElementsByClassName("conteudo");
  for (let i = 0; i < conteudos.length; i++) conteudos[i].style.display = "none";

  const botoes = document.getElementsByClassName("tablinks");
  for (let i = 0; i < botoes.length; i++) botoes[i].className = botoes[i].className.replace(" ativo", "");

  document.getElementById(nomeAba).style.display = "block";
  evt.currentTarget.className += " ativo";
}

// Abrir aba Pedidos por padrão
document.addEventListener("DOMContentLoaded", () => {
  document.querySelector(".tablinks").click();
});

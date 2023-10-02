const { JSDOM } = require('jsdom');
const fs = require('fs');


const data = fs.readFileSync('page.html', 'utf-8');
const dom = new JSDOM(data);
const { document } = dom.window;
let array = [];


for (let i = 0; document.querySelectorAll("#EdFornecedor")[i].textContent != '\nSCHERER S/A COMERCIO DE AUTOPECAS\n'; i++) {

    let obj = {
        "fornecedor": document.querySelectorAll("#EdFornecedor")[i].textContent.replace(/\n/g, ''),
        "nota": document.querySelectorAll("#EdNota")[i].textContent.replace(/\n/g, ''),
        "data": document.querySelectorAll("#EdData")[i].textContent.replace(/\n/g, ''),
        "volumes": document.querySelectorAll("#edVolumes")[i].textContent.replace(/\n/g, ''),
        "valor": document.querySelectorAll("#EdValor")[i].textContent.replace(/\n/g, '')
    }
    array.push(obj)
}


console.log("\nRelátório: \n ", array)

//////////push and post////////////

const ExcelJS = require('exceljs');
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Relatório');

worksheet.addRow(['Fornecedor', 'NF', 'Data', 'Volumes', 'Valor']);
for (let i = 0; i != array.length; i++) {

    worksheet.addRow([
        array[i].fornecedor,
        array[i].nota,
        array[i].data,
        array[i].volumes,
        array[i].valor
    ]);

}

workbook.xlsx.writeFile('planilha.xlsx')
    .then(function() {
        console.log('Planilha criada com estilos personalizados');
    })
    .catch(function(error) {
        console.log('Erro ao criar planilha:', error);
    })
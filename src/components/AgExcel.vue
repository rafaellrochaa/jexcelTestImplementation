<template>
  <div>
    <div class="mb-md">
      <input type="file" @change="carregar">
      <button @click="valores">Exibir Valores Selecionados (console)</button>
    </div>
    <div ref="displayExcel">
    </div>
  </div>
</template>

<script>
import jexcel from 'jexcel'
import 'jexcel/dist/jexcel.css'
import XLSX from 'xlsx'

export default {
  data () {
    return {
      displayExcel: null,
      posicaoSelecaoInicial: {linha: -1, coluna: -1},
      posicaoSelecaoFinal: {linha: -1, coluna: -1},
      options: {
        onselection: this.selecaoAtiva
      }
    }
  },
  methods: {
    carregarDimensoesPlanilha (workSheet) {
      let numeroLinhas = workSheet.length
      let numeroColunas = 0
      for (let row in workSheet) {
        numeroColunas = numeroColunas < workSheet[row].length? workSheet[row].length : numeroColunas
      }
      workSheet.dimensions = [numeroColunas, numeroLinhas]
    },
    carregar (fileSelected) {
      let file = fileSelected.target.files[0]
      let reader = new FileReader()
      let name = file.name
      reader.onload = e => {
        let data = e.target.result,
          fixedData = this.fixData(data),
          workbook = XLSX.read(btoa(fixedData), {type: 'base64'}),
          firstSheetName = workbook.SheetNames[0],
          worksheet = workbook.Sheets[firstSheetName]
        // let result222 = this.workbook_to_json(workbook)
        let memWorkSheet = XLSX.utils.sheet_to_json(worksheet, {header: 1, raw: false})
        this.carregarDimensoesPlanilha(memWorkSheet)
        this.options.minDimensions = memWorkSheet.dimensions
        this.options.data = memWorkSheet
        this.displayExcel = jexcel(this.$refs["displayExcel"], this.options)
      }
      reader.readAsArrayBuffer(file)
    },
    fixData (data) {
      var o = "", l = 0, w = 10240
      for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)))
      o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)))
      return o
    },
    workbook_to_json (workbook) {
      let result = {}
      workbook.SheetNames.forEach(sheetName => {
        let roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName], {header: 1, raw: false})
        if (roa.length > 0) {
          result[sheetName] = roa
        }
      })
      return result
    },
    selecaoAtiva (instance, x1, y1, x2, y2) {
      this.posicaoSelecaoInicial.linha = y1
      this.posicaoSelecaoInicial.coluna = x1
      this.posicaoSelecaoFinal.linha = y2
      this.posicaoSelecaoFinal.coluna = x2
    },
    valores () {
      let table = this.displayExcel.getData(false)
      for (let linha = this.posicaoSelecaoInicial.linha; linha <= this.posicaoSelecaoFinal.linha; linha++) {
        for (let coluna = this.posicaoSelecaoInicial.coluna; coluna <= this.posicaoSelecaoFinal.coluna; coluna++) {
          if (this.posicaoSelecaoInicial.coluna === this.posicaoSelecaoFinal.coluna) {
            if (coluna === this.posicaoSelecaoInicial.coluna && table[linha][coluna] !== '') {
              console.log('célula:', jexcel.getColumnNameFromId([coluna, linha]), 'valor: ', table[linha][coluna])
            }
          }
          else {
            if (coluna <= this.posicaoSelecaoFinal.coluna) {
              if (table[linha][coluna] !== '') {
                console.log('célula: ' + jexcel.getColumnNameFromId([coluna, linha]) + ' valor: ' + table[linha][coluna])
              }
            }
          }
        }
      }
    }
  }
}
</script>
<style>
.mb-md {
  margin-bottom: 8px;
}
</style>

<template>
  <div ref="displayExcel">
    <!-- <div><input type="button" value="Add new row" @click="() => displayExcel.insertRow()" /></div> -->
    <button @click="valores"> Exibir Valores Selecionados (console) </button>
  </div>
</template>

<script>
import jexcel from 'jexcel'
import 'jexcel/dist/jexcel.css'

export default {
  data () {
    return {
      displayExcel: null,
      posicaoSelecaoInicial: {linha: -1, coluna: -1},
      posicaoSelecaoFinal: {linha: -1, coluna: -1},
      options: {
        data: [[]],
        minDimensions: [25,30],
        onselection: this.selecaoAtiva
      }
    }
  },
  mounted () {
    this.displayExcel = jexcel(this.$refs["displayExcel"], this.options)
  },
  methods: {
    selecaoAtiva (instance, x1, y1, x2, y2, origin) {
      this.posicaoSelecaoInicial.linha = y1
      this.posicaoSelecaoInicial.coluna = x1
      this.posicaoSelecaoFinal.linha = y2
      this.posicaoSelecaoFinal.coluna = x2
    },
    valores () {
      let table = this.displayExcel.getData(false)
      for (let linha = this.posicaoSelecaoInicial.linha; linha < table.length; linha++) {
        if (linha <= this.posicaoSelecaoFinal.linha) {
          for (let coluna = 0; coluna < table[linha].length; coluna++) {
            if (this.posicaoSelecaoInicial.coluna === this.posicaoSelecaoFinal.coluna) {
              if (coluna === this.posicaoSelecaoInicial.coluna && table[linha][coluna] !== '') {
                console.log('célula:', jexcel.getColumnNameFromId([coluna, linha]), 'valor:', table[linha][coluna])
              }
            }
            else { 
              if (coluna <= this.posicaoSelecaoFinal.coluna) {
                if (table[linha][coluna] !== '') {
                  console.log('célula:', jexcel.getColumnNameFromId([coluna, linha]), 'valor:', table[linha][coluna])
                }
              }
            }
          }
        }
      }
    }
  }
}
</script>
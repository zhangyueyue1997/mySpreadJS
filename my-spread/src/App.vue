<template>
  <div>
    <div class="sample-tutorial">
      <div id="ss"
           class="sample-spreadsheets"></div>
      <div class="toolbarAll">
        <!--合并单元格-->
        <div class="toolbar">
          <div class="options-container">
            <div class="demo-options">
              <label>选择一个单元格块，然后单击下面的一个按钮.</label>
              <div class="option-row">
                <input type="button"
                       value="合并单元格"
                       title="Merge cells in selected range to one cell"
                       id="btnMerge"
                       @click="addCell" />
                <input type="button"
                       value="拆分单元格"
                       title="Unmerge cells in selected range back to all in a single cell"
                       id="btnUnMerge"
                       @click="splitCell" />
              </div>
            </div>
            <div class="demo-options"
                 style="padding-top: 16px">
              <label>选中下面的选项来打开和关闭拖动合并。</label>
              <div class="option-row">
                <label>
                  <input type="checkbox"
                         id="allowDragMerge"
                         @click="dragMerger">
                  允许拖动合并
                </label>
              </div>
            </div>
          </div>
        </div>
        <!-- 改变背景色 -->
        <div class="options-container"
             style="height: 80%;">
          <div class="option-row">
            <label>输入颜色名称来设置扩展中的灰色区域的颜色</label>
          </div>
          <div class="option-row">
            <input type="text"
                   id="grayAreaBackColor" />
            <input type="button"
                   id="setGrayAreaBackColor"
                   @click="setGrayAreaBackColor"
                   value="设置灰色区域背景色" />
          </div>
          <div class="option-row">
            <label>输入颜色名称来设置表格的背景颜色</label>
            <label class="note">注意:删除底部的背景图像可以看到这个变化</label>
          </div>
          <div class="option-row">
            <input type="text"
                   id="spreadBackColor" />
            <input type="button"
                   id="setSpreadBackColor"
                   @click="setSpreadBackColor"
                   value="设置表格背景色" />
          </div>
          <!-- <div class="option-row">
            <label>照片:</label>
            <input type="file"
                   name="image"
                   id="file_input"
                   @click="file_input" />

          </div>
          <div class="option-row">
            <label>表格背景图像布局:</label>
            <select id="layout"
                    @click="layout">
              <option value="0"
                      selected="selected">拉伸</option>
              <option value="1">中心</option>
              <option value="2">变焦</option>
              <option value="3">无</option>
            </select>
          </div>
          <div class="option-row">
            <input type="button"
                   id="setSpreadBackgroundImage"
                   @click="setSpreadBackgroundImage"
                   value="设置"
                   class="narrow-button" />
            <input type="button"
                   id="delSpreadBackgroundImage"
                   @click="delSpreadBackgroundImage"
                   value="删除背景图片" /> -->
          </div>
        </div>
      </div>
  </div>
</template>
<script>
import '@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css'
import '@grapecity/spread-sheets-vue'
import '@grapecity/spread-sheets/styles/gc.spread.sheets.excel2013white.css'
import '@grapecity/spread-sheets/dist/gc.spread.sheets.all.min.js'
// import '@grapecity/spread-sheets-resources-zh/dist/gc.spread.sheets.resources.zh.min.js'
import GC from '@grapecity/spread-sheets'

export default {
  name: 'index',
  data () {
    return {
      hostClass: 'spread-host',
      autoGenerateColumns: true,
      width: 1000,
      visible: true,
      resizable: true,
      formatter: '$ #.00',
      spread: {},
      spreadColor: {},
      pictureUrl: ''
    }
  },
  mounted () {
    // this.spread = new GC.Spread.Sheets.Workbook(document.getElementById('ss'))
    // this.initSpread(this.spread)
    this.spread = new GC.Spread.Sheets.Workbook(document.getElementById('ss'), { sheetCount: 1 })
    this.initSpreadColor(this.spread)
  },
  methods: {
    initSpread () {
      var sheet = this.spread.getSheet(0)
      sheet.suspendPaint()
      sheet.options.allowCellOverflow = true
      sheet.name('Demo')
      this.initSpreadData(sheet)
      sheet.resumePaint()
      this.addCell()
      this.splitCell()
      this.dragMerger()
    },
    dragMerger () { // 允许拖动合并
      this.spread.options.allowUserDragMerge = document.getElementById('allowDragMerge').checked
    },
    splitCell () { // 拆分单元格
      var sheet = this.spread.getActiveSheet()
      var sel = sheet.getSelections()
      if (sel.length > 0) {
        sel = this.getActualCellRange(sel[sel.length - 1], sheet.getRowCount(), sheet.getColumnCount())
        sheet.suspendPaint()
        for (var i = 0; i < sel.rowCount; i++) {
          for (var j = 0; j < sel.colCount; j++) {
            sheet.removeSpan(i + sel.row, j + sel.col)
          }
        }
        sheet.resumePaint()
      }
    },
    addCell () { // 合并单元格
      var sheet = this.spread.getActiveSheet()
      var sel = sheet.getSelections()
      if (sel.length > 0) {
        sel = this.getActualCellRange(sel[sel.length - 1], sheet.getRowCount(), sheet.getColumnCount())
        sheet.addSpan(sel.row, sel.col, sel.rowCount, sel.colCount)
      }
    },
    getActualCellRange (cellRange, rowCount, columnCount) {
      var spreadNS = GC.Spread.Sheets
      // eslint-disable-next-line eqeqeq
      if (cellRange.row == -1 && cellRange.col == -1) {
        return new spreadNS.Range(0, 0, rowCount, columnCount)
        // eslint-disable-next-line eqeqeq
      } else if (cellRange.row == -1) {
        return new spreadNS.Range(0, cellRange.col, rowCount, cellRange.colCount)
        // eslint-disable-next-line eqeqeq
      } else if (cellRange.col == -1) {
        return new spreadNS.Range(cellRange.row, 0, cellRange.rowCount, columnCount)
      }
      return cellRange
    },
    initSpreadData (sheet) {
      var spreadNS = GC.Spread.Sheets
      sheet.addSpan(1, 1, 1, 3)
      sheet.setValue(1, 1, 'Store')
      sheet.addSpan(1, 4, 1, 7)
      sheet.setValue(1, 4, 'Goods')
      sheet.addSpan(2, 1, 1, 2)
      sheet.setValue(2, 1, 'Area')
      sheet.addSpan(2, 3, 2, 1)
      sheet.setValue(2, 3, 'ID')
      sheet.addSpan(2, 4, 1, 2)
      sheet.setValue(2, 4, 'Fruits')
      sheet.addSpan(2, 6, 1, 2)
      sheet.setValue(2, 6, 'Vegetables')
      sheet.addSpan(2, 8, 1, 2)
      sheet.setValue(2, 8, 'Foods')
      sheet.addSpan(2, 10, 2, 1)
      sheet.setValue(2, 10, 'Total')

      sheet.setValue(3, 1, 'State')
      sheet.setValue(3, 2, 'City')
      sheet.setValue(3, 4, 'Grape')
      sheet.setValue(3, 5, 'Apple')
      sheet.setValue(3, 6, 'Potato')
      sheet.setValue(3, 7, 'Tomato')
      sheet.setValue(3, 8, 'SandWich')
      sheet.setValue(3, 9, 'Hamburger')

      sheet.addSpan(4, 1, 7, 1)
      sheet.addSpan(4, 2, 3, 1)
      sheet.addSpan(7, 2, 3, 1)
      sheet.addSpan(10, 2, 1, 2)
      sheet.setValue(10, 2, 'Sub Total:')
      sheet.addSpan(11, 1, 7, 1)
      sheet.addSpan(11, 2, 3, 1)
      sheet.addSpan(14, 2, 3, 1)
      sheet.addSpan(17, 2, 1, 2)
      sheet.setValue(17, 2, 'Sub Total:')
      sheet.addSpan(18, 1, 1, 3)
      sheet.setValue(18, 1, 'Total:')

      sheet.setValue(4, 1, 'NC')
      sheet.setValue(4, 2, 'Raleigh')
      sheet.setValue(7, 2, 'Charlotte')
      sheet.setValue(4, 3, '001')
      sheet.setValue(5, 3, '002')
      sheet.setValue(6, 3, '003')
      sheet.setValue(7, 3, '004')
      sheet.setValue(8, 3, '005')
      sheet.setValue(9, 3, '006')
      sheet.setValue(11, 1, 'PA')
      sheet.setValue(11, 2, 'Philadelphia')
      sheet.setValue(14, 2, 'Pittsburgh')
      sheet.setValue(11, 3, '007')
      sheet.setValue(12, 3, '008')
      sheet.setValue(13, 3, '009')
      sheet.setValue(14, 3, '010')
      sheet.setValue(15, 3, '011')
      sheet.setValue(16, 3, '012')

      sheet.setFormula(10, 4, '=SUM(E5:E10)')
      sheet.setFormula(10, 5, '=SUM(F5:F10)')
      sheet.setFormula(10, 6, '=SUM(G5:G10)')
      sheet.setFormula(10, 7, '=SUM(H5:H10)')
      sheet.setFormula(10, 8, '=SUM(I5:I10)')
      sheet.setFormula(10, 9, '=SUM(J5:J10)')

      sheet.setFormula(17, 4, '=SUM(E12:E17)')
      sheet.setFormula(17, 5, '=SUM(F12:F17)')
      sheet.setFormula(17, 6, '=SUM(G12:G17)')
      sheet.setFormula(17, 7, '=SUM(H12:H17)')
      sheet.setFormula(17, 8, '=SUM(I12:I17)')
      sheet.setFormula(17, 9, '=SUM(J12:J17)')

      for (var i = 0; i < 14; i++) {
        sheet.setFormula(4 + i, 10, '=SUM(E' + (5 + i).toString() + ':J' + (5 + i).toString() + ')')
      }

      sheet.setFormula(18, 4, '=E11+E18')
      sheet.setFormula(18, 5, '=F11+F18')
      sheet.setFormula(18, 6, '=G11+G18')
      sheet.setFormula(18, 7, '=H11+H18')
      sheet.setFormula(18, 8, '=I11+I18')
      sheet.setFormula(18, 9, '=J11+J18')
      sheet.setFormula(18, 10, '=K11+K18')

      sheet.getRange(1, 1, 3, 10).backColor('#D9D9FF')
      sheet.getRange(4, 1, 15, 3).backColor('#D9FFD9')
      sheet.getRange(1, 1, 3, 10).hAlign(spreadNS.HorizontalAlign.center)

      sheet.getRange(1, 1, 18, 10).setBorder(new spreadNS.LineBorder('Black', spreadNS.LineStyle.thin), { all: true })
      sheet.getRange(4, 4, 3, 6).setBorder(new spreadNS.LineBorder('Black', spreadNS.LineStyle.dotted), { inside: true })
      sheet.getRange(7, 4, 3, 6).setBorder(new spreadNS.LineBorder('Black', spreadNS.LineStyle.dotted), { inside: true })
      sheet
        .getRange(11, 4, 3, 6)
        .setBorder(new spreadNS.LineBorder('Black', spreadNS.LineStyle.dotted), { inside: true })
      sheet
        .getRange(14, 4, 3, 6)
        .setBorder(new spreadNS.LineBorder('Black', spreadNS.LineStyle.dotted), { inside: true })

      this.fillSampleData(sheet, new spreadNS.Range(4, 4, 6, 6))
      this.fillSampleData(sheet, new spreadNS.Range(11, 4, 6, 6))
    },
    fillSampleData (sheet, range) {
      for (var i = 0; i < range.rowCount; i++) {
        for (var j = 0; j < range.colCount; j++) {
          sheet.setValue(range.row + i, range.col + j, Math.ceil(Math.random() * 300))
        }
      }
    },

    // 改变背景色
    initSpreadColor () {
      var spreadNS = GC.Spread.Sheets
      var sheet = this.spread.getSheet(0)
      this.pictureUrl = ''
      this.spread.suspendPaint()
      sheet.setRowCount(20)
      sheet.setColumnCount(20)
      this.spread.options.backColor = 'white'
      this.spread.options.grayAreaBackColor = 'gray'
      this.spread.options.backgroundImageLayout = spreadNS.ImageLayout.stretch
      // this.spreadColor.options.backgroundImage = 'https://demo.grapecity.com.cn/spreadjs/SpreadJSTutorial/spread/source/images/backImage.png'

      this.spread.resumePaint()
      this.layout()
      this.setGrayAreaBackColor()
      this.setSpreadBackColor()
      this.file_input()
      this.setSpreadBackgroundImage()
      this.delSpreadBackgroundImage()
    },
    layout () {
      var layout = parseInt(document.getElementById('layout').value)
      this.spread.options.backgroundImageLayout = layout
    },
    setGrayAreaBackColor () {
      var color = document.getElementById('grayAreaBackColor').value
      this.spread.options.grayAreaBackColor = color
    },
    _getElementById (id) {
      return document.getElementById(id)
    },
    setSpreadBackColor () {
      var color = document.getElementById('spreadBackColor').value
      this.spread.options.backColor = color
    },
    file_input () {
      var file = this.files[0]
      if (!/image\/\w+/.test(file.type)) {
        alert('The file muse be image!')
        return false
      }
      var reader = new FileReader()
      reader.readAsDataURL(file)
      reader.onload = function (e) {
        this.pictureUrl = this.result
      }
    },
    setSpreadBackgroundImage () {
      this.spread.options.backgroundImage = this.pictureUrl
    },
    delSpreadBackgroundImage () {
      this.spread.options.backgroundImage = ''
    }
  }
}
</script>
<style scoped>
#ss {
  width: 100%;
  height: 700px;
  border: 1px solid gray;
}
.spread-host {
  width: 100%;
  height: 700px;
  display: flex;
}

.sample-tutorial {
  position: relative;
  height: 100%;
  overflow: hidden;
  font-size: 14px;
  display: flex;
}

.sample-spreadsheets {
  width: calc(100% - 280px);
  height: 100%;
  overflow: hidden;
  float: left;
}

.options-container {
  float: right;
  width: 280px;
  padding: 12px;
  height: 100%;
  box-sizing: border-box;
  background: #fbfbfb;
  overflow: auto;
}

.option-row {
  font-size: 14px;
  padding: 5px;
  margin-top: 10px;
}

label {
  margin-bottom: 6px;
}

input {
  padding: 4px 6px;
}

input[type="button"] {
  margin-top: 6px;
}
body {
  position: absolute;
  top: 0;
  bottom: 0;
  left: 0;
  right: 0;
}
.toolbar {
  display: flex;
  flex-direction: column;
}
.buttonColor {
  padding: 5px;
  background-color: #b5dd63;
  /* border-radius: 5px; */
  margin-top: 50px;
  font-weight: 700;
}
.toolbarAll {
  display: flex;
  flex-direction: column;
  padding: 10px;
}
</style>

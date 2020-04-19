<template>
  <div class="sample-tutorial"
       style="height: 76%;">
    <div id="ss"
         class="sample-spreadsheets"></div>
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
      <div class="option-row">
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
               value="删除背景图片" />
      </div>
    </div>
  </div>
</template>
<script>
import '@grapecity/spread-sheets/styles/gc.spread.sheets.excel2013white.css'
import '@grapecity/spread-sheets/dist/gc.spread.sheets.all.min.js'
import '@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css'
import '@grapecity/spread-sheets-vue'
import GC from '@grapecity/spread-sheets'
export default {
  name: 'backgroundColor',
  components: {
  },
  data () {
    return {
      //   hostClass: 'spread-host',
      //   autoGenerateColumns: true,
      //   width: 1000,
      //   visible: true,
      //   resizable: true,
      //   formatter: '$ #.00',
      spread: {},
      pictureUrl: ''
    }
  },
  mounted () {
    this.spread = new GC.Spread.Sheets.Workbook(document.getElementById('ss'), { sheetCount: 1 })
    this.initSpread(this.spread)
  },
  methods: {
    initSpread () {
      var spreadNS = GC.Spread.Sheets
      var sheet = this.spread.getSheet(0)
      this.pictureUrl = ''
      this.spread.suspendPaint()
      sheet.setRowCount(15)
      sheet.setColumnCount(20)
      this.spread.options.backColor = 'white'
      this.spread.options.grayAreaBackColor = 'gray'
      this.spread.options.backgroundImageLayout = spreadNS.ImageLayout.stretch
      this.spread.options.backgroundImage = 'https://demo.grapecity.com.cn/spreadjs/SpreadJSTutorial/spread/source/images/backImage.png'

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
<style  scoped>
input[type="button"] {
  width: 180px;
}

input[type="text"] {
  padding: 4px;
  margin-top: 4px;
  width: 100%;
  box-sizing: border-box;
}

label {
  /*不同 */
  display: inline-block;
  margin-bottom: 6px;
}
.note {
  margin-top: 0px;
}
.colorLabel {
  background-color: lavender;
}

.wide-label {
  width: 260px;
}

input.narrow-button,
.narrow-label {
  width: 74px;
}
.sample-tutorial {
  position: relative;
  height: 100%;
  overflow: hidden;
}

.sample-spreadsheets {
  width: calc(100% - 300px);
  height: 100%;
  overflow: hidden;
  float: left;
}

.options-container {
  float: right;
  width: 300px;
  padding: 12px;
  height: 100%;
  box-sizing: border-box;
  background: #fbfbfb;
  overflow: auto;
}
.option-row:nth-child(1) {
  padding-bottom: 0px;
}
.option-row:nth-child(2) {
  margin-top: 0px;
  padding-top: 0px;
}
.option-row:nth-child(3) {
  padding-bottom: 0px;
}

.option-row:nth-child(4) {
  margin-top: 0px;
  padding-top: 0px;
}
.option-row {
  font-size: 14px;
  padding: 5px;
  margin-top: 10px;
}

input {
  padding: 4px 6px;
}

input[type="button"] {
  margin-top: 6px;
  display: block;
}

hr {
  border-color: #fff;
  opacity: 0.2;
  margin-top: 20px;
}

body {
  position: absolute;
  top: 0;
  bottom: 0;
  left: 0;
  right: 0;
}
</style>

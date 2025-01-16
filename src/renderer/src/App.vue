<template>
  <div>
    <h1>基准DBC</h1>
    <div class="container">
      <div class="item">
        <div class="drag-box">请将文件拖入此处或点击选择文件</div>
        <input
            type="file"
            accept=".xls,.xlsx"
            class="upload_file"
            @change="readBaseDbc($event)"
        />
      </div>
      <div class="item">
        <div><span>报文数：</span><span>{{ baseMsgCount }} ( {{ baseMsgMinIndex }} ~ {{ baseMsgMaxIndex }} )</span></div>
        <div><span>信号数：</span><span>{{ baseSignalCount }} ( {{ baseSignalMinIndex }} ~ {{
            baseSignalMaxIndex
          }} )</span></div>
      </div>
    </div>
    <h1>新版DBC</h1>
    <div class="container">
      <div class="item">
        <div class="drag-box">请将文件拖入此处或点击选择文件</div>
        <input
            type="file"
            accept=".xls,.xlsx"
            class="upload_file"
            @click="checkBaseDbc($event)"
            @change="readNewDbc($event)"
        />
      </div>
      <div class="item">
        <div><span>总报文数：</span><span>{{ dbcMsgCount }}</span></div>
        <div><span>匹配报文数：</span><span>{{ dbcMsgMatchCount }}</span></div>
        <div><span>新增报文数：</span><span>{{ dbcMsgNewCount }} ( {{ dbcMsgNewMinIndex }} ~ {{
            dbcMsgNewMaxIndex
          }} )</span></div>
        <div><span>总信号数：</span><span>{{ dbcSignalCount }}</span></div>
        <div><span>匹配信号数：</span><span>{{ dbcSignalMatchCount }}</span></div>
        <div><span>新增信号数：</span><span>{{ dbcSignalNewCount }} ( {{
            dbcSignalNewMinIndex
          }} ~ {{ dbcSignalNewMaxIndex }} )</span></div>
      </div>
    </div>
    <div class="function-button">
      <input type="button" style="margin-left: 20px; margin-right: 20px;" value="导出重排后新版DBC"
             @click="exportNewDbc($event)"/>
      <input type="button" style="margin-left: 20px; margin-right: 20px;" value="导出汇总后基准DBC"
             @click="exportBaseDbc($event)"/>
      <input type="button" style="margin-left: 20px; margin-right: 20px;" value="导出差异DBC"
             @click="exportDiffDbc($event)"/>
    </div>
  </div>
  <div>

  </div>
</template>

<script>
import * as XLSX from "xlsx"

export default {
  data() {
    return {
      // 基础DBC头
      baseDbcHeaderMap: new Map(),
      // 新版DBC头
      newDbcHeaderMap: new Map(),
      // 基准报文索引Map
      baseMsgIndexMap: new Map(),
      // 基准报文数量
      baseMsgCount: 0,
      // 基准报文最小索引
      baseMsgMinIndex: 0,
      // 基准报文最大索引
      baseMsgMaxIndex: 0,
      // 基准信号Map
      baseSignalMap: new Map(),
      // 基准信号数量
      baseSignalCount: 0,
      // 基准信号最小索引
      baseSignalMinIndex: 0,
      // 基准信号最大索引
      baseSignalMaxIndex: 0,
      // DBC报文Map
      dbcMsgMap: new Map(),
      // DBC报文数量
      dbcMsgCount: 0,
      // DBC报文匹配基准报文Map
      dbcMsgMatchMap: new Map(),
      // DBC报文匹配基准报文数量
      dbcMsgMatchCount: 0,
      // DBC新增报文索引Map
      dbcMsgNewIndexMap: new Map(),
      // DBC新增报文数量
      dbcMsgNewCount: 0,
      // DBC新增报文最小索引
      dbcMsgNewMinIndex: 0,
      // DBC新增报文最大索引
      dbcMsgNewMaxIndex: 0,
      // DBC信号Map
      dbcSignalMap: new Map(),
      // DBC信号数量
      dbcSignalCount: 0,
      // DBC信号匹配基准信号Map
      dbcSignalMatchMap: new Map(),
      // DBC信号匹配基准信号数量
      dbcSignalMatchCount: 0,
      // DBC新增信号索引Map
      dbcSignalNewIndexMap: new Map(),
      // DBC新增信号数量
      dbcSignalNewCount: 0,
      // DBC新增信号最小索引
      dbcSignalNewMinIndex: 0,
      // DBC新增信号最大索引
      dbcSignalNewMaxIndex: 0,
      // DBC报文比基础版本缺少的报文Map
      dbcMsgDiffMap: new Map(),
      // DBC信号比基础版本缺少的信号Map
      dbcSignalDiffMap: new Map(),
      lists: []
    }
  },
  methods: {
    // 读取基准DBC
    readBaseDbc: async function (evt) {
      let files = evt.target.files;
      if (!files || files.length === 0) return;
      let file = files[0];

      new Promise((resolve) => {
        let reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = (ev) => {
          resolve(ev.target.result);
        }
      }).then((dataBinary) => {
        let workBook = XLSX.read(dataBinary, {type: "binary", cellDates: true});
        let firstWorkSheet = workBook.Sheets[workBook.SheetNames[0]];
        const header = this.getHeaderRow(firstWorkSheet);
        this.parseBaseDbcHeader(header);
        const data = XLSX.utils.sheet_to_json(firstWorkSheet);
        // console.log("读取所有excel数据", data);
        let lastMsgKey;
        data.forEach((item) => {
          let index;
          if (this.isBaseMsgRow(item)) {
            let msgKey = this.getBaseMsgKey(item);
            index = this.getBaseKey(item, "msg_index");
            this.baseMsgIndexMap.set(msgKey, index)
            if (this.baseMsgMinIndex === 0 || index < this.baseMsgMinIndex) {
              this.baseMsgMinIndex = index;
            }
            if (this.baseMsgMaxIndex === 0 || index > this.baseMsgMaxIndex) {
              this.baseMsgMaxIndex = index;
            }
            lastMsgKey = msgKey
          } else if (this.isBaseSignalRow(item)) {
            let signalKey = this.getBaseSignalKey(lastMsgKey, item);
            index = this.getBaseKey(item, "signal_index");
            this.baseSignalMap.set(signalKey, index)
            if (this.baseSignalMinIndex === 0 || index < this.baseSignalMinIndex) {
              this.baseSignalMinIndex = index;
            }
            if (this.baseSignalMaxIndex === 0 || index > this.baseSignalMaxIndex) {
              this.baseSignalMaxIndex = index;
            }
          }
        });
        this.baseMsgCount = this.baseMsgIndexMap.size;
        this.baseSignalCount = this.baseSignalMap.size;
      });
    },
    // 解析基准DBC头
    parseBaseDbcHeader: function (row) {
      row.forEach((header) => {
        // 解析报文名称
        if (header.toLowerCase().includes("msg_name")
            || header.toLowerCase().includes("msg name")
            || header.toLowerCase().includes("报文名称")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.baseDbcHeaderMap.set("msg_name", header);
        }
        // 解析报文ID
        if (header.toLowerCase().includes("msg_id")
            || header.toLowerCase().includes("msg id")
            || header.toLowerCase().includes("报文标识符")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.baseDbcHeaderMap.set("msg_id", header);
        }
        // 解析报文长度
        if (header.toLowerCase().includes("msg_length")
            || header.toLowerCase().includes("msg length")
            || header.toLowerCase().includes("报文长度")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.baseDbcHeaderMap.set("msg_length", header);
        }
        // 解析报文索引
        if (header.toLowerCase().includes("re_message_id")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.baseDbcHeaderMap.set("msg_index", header);
        }
        // 解析信号名称
        if (header.toLowerCase().includes("signal_name")
            || header.toLowerCase().includes("signal name")
            || header.toLowerCase().includes("信号名称")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.baseDbcHeaderMap.set("signal_name", header);
        }
        // 解析信号索引
        if (header.toLowerCase().includes("re_signal_id")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.baseDbcHeaderMap.set("signal_index", header);
        }
      });
      console.log(this.baseDbcHeaderMap);
    },
    // 是否是基本报文行
    isBaseMsgRow: function (row) {
      return this.getBaseKey(row, "msg_name") != null || this.getBaseKey(row, "msg_index") != null;
    },
    // 是否是新版报文行
    isNewMsgRow: function (row) {
      return this.getNewKey(row, "msg_name") != null;
    },
    // 是否是基准信号行
    isBaseSignalRow: function (row) {
      return this.getBaseKey(row, "signal_name") != null || this.getBaseKey(row, "signal_index") != null;
    },
    // 是否是新版信号行
    isNewSignalRow: function (row) {
      return this.getNewKey(row, "signal_name") != null || this.getNewKey(row, "signal_index") != null;
    },
    // 获取基准报文Key
    getBaseMsgKey: function (row) {
      return this.getBaseKey(row, "msg_name") + "-" + this.getBaseKey(row, "msg_id") + "-" + this.getBaseKey(row, "msg_length");
    },
    // 获取新版报文Key
    getNewMsgKey: function (row) {
      return this.getNewKey(row, "msg_name") + "-" + this.getNewKey(row, "msg_id") + "-" + this.getNewKey(row, "msg_length");
    },
    // 获取基准信号Key
    getBaseSignalKey: function (msgKey, row) {
      return msgKey + "-" + this.getBaseKey(row, "signal_name")
    },
    // 获取新版信号Key
    getNewSignalKey: function (msgKey, row) {
      return msgKey + "-" + this.getNewKey(row, "signal_name")
    },
    // 检查基准DBC是否加载
    checkBaseDbc: function (evt) {
      if (this.baseMsgCount === 0) {
        evt.preventDefault();
        alert("请先选择基准DBC");
        return false;
      }
      return true;
    },
    // 检查新版DBC是否加载
    checkNewDbc: function (evt) {
      if (this.dbcMsgCount === 0) {
        evt.preventDefault();
        alert("请先选择新版DBC");
        return false;
      }
      return true;
    },
    // 读取新版DBC
    readNewDbc: async function (evt) {
      let files = evt.target.files;
      if (!files || files.length === 0) return;
      let file = files[0];

      new Promise((resolve) => {
        let reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = (ev) => {
          resolve(ev.target.result);
        }
      }).then((dataBinary) => {
        let workBook = XLSX.read(dataBinary, {type: "binary", cellDates: true});
        workBook.SheetNames.forEach((sheetName) => {
          let currentSheet = workBook.Sheets[sheetName];
          // console.log("sheet", currentSheet)
          if (!this.isSignalSheet(currentSheet)) {
            console.log("跳过sheet", sheetName)
            return;
          }
          const header = this.getHeaderRow(currentSheet);
          this.parseNewDbcHeader(header);
          const data = XLSX.utils.sheet_to_json(currentSheet);
          // console.log("读取所有excel数据", data);
          let lastMsgKey;
          data.forEach((item) => {
            if (this.isNewMsgRow(item)) {
              let msgKey = this.getNewMsgKey(item);
              if (this.getNewKey(item, "source_node") == null && this.getNewKey(item, "msg_type") === "Diag") {
                item["源节点"] = "RIDS"
              }
              this.dbcMsgMap.set(msgKey, item);
              if (this.baseMsgIndexMap.has(msgKey)) {
                this.dbcMsgMatchMap.set(msgKey, this.baseMsgIndexMap.get(msgKey));
              } else {
                if (this.dbcMsgNewMinIndex === 0) {
                  this.dbcMsgNewMinIndex = this.baseMsgMaxIndex + 1;
                  this.dbcMsgNewMaxIndex = this.baseMsgMaxIndex;
                }
                if (!this.dbcMsgNewIndexMap.has(msgKey)) {
                  this.dbcMsgNewMaxIndex++
                  this.dbcMsgNewIndexMap.set(msgKey, this.dbcMsgNewMaxIndex);
                }
              }
              lastMsgKey = msgKey
            } else if (this.isNewSignalRow(item)) {
              var signalKey = this.getNewSignalKey(lastMsgKey, item)
              // 如果不存在源节点并且是诊断类型则默认源节点设置为诊断
              if (this.getNewKey(item, "source_node") == null && this.getNewKey(this.dbcMsgMap.get(lastMsgKey), "msg_type") === "Diag") {
                item["源节点"] = "RIDS"
              }
              if (!this.dbcSignalMap.has(signalKey)) {
                this.dbcSignalMap.set(signalKey, item);
              }
              if (this.baseSignalMap.has(signalKey)) {
                this.dbcSignalMatchMap.set(signalKey, this.baseSignalMap.get(signalKey));
              } else {
                if (this.dbcSignalNewMinIndex === 0) {
                  this.dbcSignalNewMinIndex = this.baseSignalMaxIndex + 1;
                  this.dbcSignalNewMaxIndex = this.baseSignalMaxIndex;
                }
                if (!this.dbcSignalNewIndexMap.has(signalKey)) {
                  this.dbcSignalNewMaxIndex++
                  this.dbcSignalNewIndexMap.set(signalKey, this.dbcSignalNewMaxIndex);
                  // console.log("不匹配", signalKey)
                }
              }
            } else {
              console.log("该行异常", item)
            }
          });
          this.dbcMsgCount = this.dbcMsgMap.size;
          this.dbcMsgMatchCount = this.dbcMsgMatchMap.size;
          this.dbcMsgNewCount = this.dbcMsgNewIndexMap.size;
          this.dbcSignalCount = this.dbcSignalMap.size;
          this.dbcSignalMatchCount = this.dbcSignalMatchMap.size;
          this.dbcSignalNewCount = this.dbcSignalNewIndexMap.size;
        })
      });
    },
    // 解析新版DBC头
    parseNewDbcHeader: function (row) {
      row.forEach((header) => {
        // 解析源节点
        if (header.toLowerCase().includes("source_node")
            || header.toLowerCase().includes("source node")
            || header.toLowerCase().includes("源节点")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("source_node", header);
        }
        // 解析路由
        if (header.toLowerCase().includes("routing")
            || header.toLowerCase().includes("路由")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("routing", header);
        }
        // 解析报文名称
        if (header.toLowerCase().includes("msg_name")
            || header.toLowerCase().includes("msg name")
            || header.toLowerCase().includes("报文名称")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("msg_name", header);
        }
        // 解析报文ID
        if (header.toLowerCase().includes("msg_id")
            || header.toLowerCase().includes("msg id")
            || header.toLowerCase().includes("报文标识符")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("msg_id", header);
        }
        // 解析报文长度
        if (header.toLowerCase().includes("msg_length")
            || header.toLowerCase().includes("msg length")
            || header.toLowerCase().includes("报文长度")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("msg_length", header);
        }
        // 解析报文周期时间
        if (header.toLowerCase().includes("msg_cycle_time")
            || header.toLowerCase().includes("msg cycle time (ms)")
            || header.toLowerCase().includes("报文周期时间")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("msg_cycle_time", header);
        }
        // 解析报文发送的快速周期
        if (header.toLowerCase().includes("msg_cycle_time_fast")
            || header.toLowerCase().includes("msg cycle time fast")
            || header.toLowerCase().includes("报文发送的快速周期")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("msg_cycle_time_fast", header);
        }
        // 解析报文快速发送的次数
        if (header.toLowerCase().includes("msg_nr_of_repetition")
            || header.toLowerCase().includes("msg nr. of repetition")
            || header.toLowerCase().includes("报文快速发送的次数")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("msg_nr_of_repetition", header);
        }
        // 解析报文延时时间
        if (header.toLowerCase().includes("msg_delay_time")
            || header.toLowerCase().includes("msg delay time")
            || header.toLowerCase().includes("报文快速发送的次数")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("msg_delay_time", header);
        }
        // 解析报文类型
        if (header.toLowerCase().includes("msg_type")
            || header.toLowerCase().includes("msg type")
            || header.toLowerCase().includes("报文类型")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("msg_type", header);
        }
        // 解析报文发送类型
        if (header.toLowerCase().includes("msg_send_type")
            || header.toLowerCase().includes("msg send type")
            || header.toLowerCase().includes("报文发送类型")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("msg_send_type", header);
        }
        // 解析信号名称
        if (header.toLowerCase().includes("signal_name")
            || header.toLowerCase().includes("signal name")
            || header.toLowerCase().includes("信号名称")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_name", header);
        }
        // 解析信号描述
        if (header.toLowerCase().includes("signal_description")
            || header.toLowerCase().includes("signal description")
            || header.toLowerCase().includes("信号描述")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_description", header);
        }
        // 解析信号排列格式
        if (header.toLowerCase().includes("byte_order")
            || header.toLowerCase().includes("byte order")
            || header.toLowerCase().includes("排列格式")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_byte_order", header);
        }
        // 解析信号起始字节
        if (header.toLowerCase().includes("start_byte")
            || header.toLowerCase().includes("start byte")
            || header.toLowerCase().includes("起始字节")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_start_byte", header);
        }
        // 解析信号起始位
        if (header.toLowerCase().includes("start_bit")
            || header.toLowerCase().includes("start bit")
            || header.toLowerCase().includes("起始位")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_start_bit", header);
        }
        // 解析信号发送类型
        if (header.toLowerCase().includes("signal_send_type")
            || header.toLowerCase().includes("signal send type")
            || header.toLowerCase().includes("信号发送类型")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_send_type", header);
        }
        // 解析信号长度
        if (header.toLowerCase().includes("bit_length")
            || header.toLowerCase().includes("bit length")
            || header.toLowerCase().includes("信号长度")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_bit_length", header);
        }
        // 解析信号数据类型
        if (header.toLowerCase().includes("data_type")
            || header.toLowerCase().includes("data type")
            || header.toLowerCase().includes("信号数据类型")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_data_type", header);
        }
        // 解析信号精度
        if (header.toLowerCase().includes("resolution")
            || header.toLowerCase().includes("精度")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_resolution", header);
        }
        // 解析信号偏移量
        if (header.toLowerCase().includes("offset")
            || header.toLowerCase().includes("偏移量")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_offset", header);
        }
        // 解析信号物理最小值
        if (header.toLowerCase().includes("signal_min_value_phys")
            || header.toLowerCase().includes("signal min. value (phys)")
            || header.toLowerCase().includes("物理最小值")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_min_value_phys", header);
        }
        // 解析信号物理最大值
        if (header.toLowerCase().includes("signal_max_value_phys")
            || header.toLowerCase().includes("signal max. value(phys)")
            || header.toLowerCase().includes("物理最大值")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_max_value_phys", header);
        }
        // 解析信号总线最小值
        if (header.toLowerCase().includes("signal_min_value_hex")
            || header.toLowerCase().includes("signal min. value (hex)")
            || header.toLowerCase().includes("总线最小值")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_min_value_hex", header);
        }
        // 解析信号总线最大值
        if (header.toLowerCase().includes("signal_max_value_hex")
            || header.toLowerCase().includes("signal max. value(hex)")
            || header.toLowerCase().includes("总线最大值")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_max_value_hex", header);
        }
        // 解析信号初始值
        if (header.toLowerCase().includes("initial_value_hex")
            || header.toLowerCase().includes("initial value(hex)")
            || header.toLowerCase().includes("初始值")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_initial_value_hex", header);
        }
        // 解析信号无效值
        if (header.toLowerCase().includes("invalid_value_hex")
            || header.toLowerCase().includes("invalid value(hex)")
            || header.toLowerCase().includes("无效值")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_invalid_value_hex", header);
        }
        // 解析信号非使能值
        if (header.toLowerCase().includes("inactive_value_hex")
            || header.toLowerCase().includes("inactive value(hex)")
            || header.toLowerCase().includes("非使能值")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_inactive_value_hex", header);
        }
        // 解析信号单位
        if (header.toLowerCase().includes("unit")
            || header.toLowerCase().includes("单位")) {
          this.newDbcHeaderMap.set("signal_unit", header);
        }
        // 解析信号信号值描述
        if (header.toLowerCase().includes("signal_value_description")
            || header.toLowerCase().includes("signal value description")
            || header.toLowerCase().includes("信号值描述")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_value_description", header);
        }
        // 解析信号索引
        if (header.toLowerCase().includes("re_signal_id")) {
          header = header.trim();
          header = header.replace(/[\r\n]/g, "");
          header = header.replace(" ", "");
          this.newDbcHeaderMap.set("signal_index", header);
        }
      });
      // console.log(this.newDbcHeaderMap);
    },
    // 是否是信号Sheet
    isSignalSheet: function (sheet) {
      const header = this.getHeaderRow(sheet);
      // console.log("读取的excel表头数据（第一行）", header);
      return header[0].startsWith("Msg")
    },
    // 导出基准DBC
    exportBaseDbc(evt) {
      if (!this.checkBaseDbc(evt)) {
        return;
      }
      if (!this.checkNewDbc(evt)) {
        return;
      }
      let newBaseDbc = [];
      let baseMsgIndexSortedMap = Array.from(this.baseMsgIndexMap);
      baseMsgIndexSortedMap.sort(function (a, b) {
        return a[0].localeCompare(b[0])
      });
      for (let [msgKey, msgValue] of baseMsgIndexSortedMap) {
        // baseMsgIndexSortedMap.forEach((msgValue, msgKey) => {
        const newMsg = [{
          "msg_name": msgKey.split("-")[0],
          "msg_id": msgKey.split("-")[1],
          "msg_length": msgKey.split("-")[2],
          "re_message_id": msgValue
        }]
        newBaseDbc = newBaseDbc.concat(newMsg)
        this.baseSignalMap.forEach((signalValue, signalKey) => {
          if (signalKey.startsWith(msgKey)) {
            const newSignal = [{
              "signal_name": signalKey.split("-")[3],
              "re_signal_id": signalValue
            }]
            newBaseDbc = newBaseDbc.concat(newSignal)
          }
        })
        this.dbcSignalNewIndexMap.forEach((signalValue, signalKey) => {
          if (signalKey.startsWith(msgKey)) {
            const newSignal = [{
              "signal_name": signalKey.split("-")[3],
              "re_signal_id": signalValue
            }]
            newBaseDbc = newBaseDbc.concat(newSignal)
          }
        })
      }
      let dbcMsgNewIndexSortedMap = Array.from(this.dbcMsgNewIndexMap);
      dbcMsgNewIndexSortedMap.sort(function (a, b) {
        return a[0].localeCompare(b[0])
      });
      for (let [msgKey, msgValue] of dbcMsgNewIndexSortedMap) {
        // dbcMsgNewIndexSortedMap.forEach((msgValue, msgKey) => {
        if (!this.baseMsgIndexMap.has(msgKey)) {
          const newMsg = [{
            "msg_name": msgKey.split("-")[0],
            "msg_id": msgKey.split("-")[1],
            "msg_length": msgKey.split("-")[2],
            "re_message_id": msgValue
          }]
          newBaseDbc = newBaseDbc.concat(newMsg)
          this.dbcSignalNewIndexMap.forEach((signalValue, signalKey) => {
            if (signalKey.startsWith(msgKey)) {
              const newSignal = [{
                "signal_name": signalKey.split("-")[3],
                "re_signal_id": signalValue
              }]
              newBaseDbc = newBaseDbc.concat(newSignal)
            }
          })
        }
      }
      const jsonWorkSheet = XLSX.utils.json_to_sheet(newBaseDbc)
      const workBook = {
        SheetNames: ["total"],
        Sheets: {
          ["total"]: jsonWorkSheet,
        }
      };
      const today = new Date()
      const date = today.getFullYear() + "" + (today.getMonth() + 1).toString().padStart(2, "0") + "" + today.getDate().toString().padStart(2, "0")
      const fileName = "基准DBC_" + date + ".xlsx"
      XLSX.writeFile(workBook, fileName);
    },
    // 导出新版DBC
    exportNewDbc(evt) {
      if (!this.checkBaseDbc(evt)) {
        return;
      }
      if (!this.checkNewDbc(evt)) {
        return;
      }
      let newNewDbc = [];
      let dbcMsgSortedMap = Array.from(this.dbcMsgMap);
      dbcMsgSortedMap.sort(function (a, b) {
        return a[0].localeCompare(b[0])
      });
      for (let [msgKey, msgValue] of dbcMsgSortedMap) {
        // dbcMsgSortedMap.forEach((msgValue, msgKey) => {
        // msgKey = this.getNewMsgKey(msgValue);
        let msgIndex = this.baseMsgIndexMap.get(msgKey);
        let msgNew = false;
        if (msgIndex == null) {
          msgIndex = this.dbcMsgNewIndexMap.get(msgKey)
          if (msgIndex == null) {
            console.log("msgKey", msgKey)
          } else {
            msgNew = true
          }
        }
        const newMsg = [{
          "Msg Name\n报文名称": this.getNewKey(msgValue, "msg_name"),
          "Msg Type\n报文类型": this.getNewKey(msgValue, "msg_type"),
          "Msg ID\n报文标识符": this.getNewKey(msgValue, "msg_id"),
          "Msg Send Type\n报文发送类型": this.getNewKey(msgValue, "msg_send_type"),
          "Msg Cycle Time (ms)\n报文周期时间": this.getNewKey(msgValue, "msg_cycle_time"),
          "Msg Length (Byte)\n报文长度": this.getNewKey(msgValue, "msg_length"),
          "Signal Name\n信号名称": "",
          "Signal Description\n信号描述": this.getNewKey(msgValue, "signal_description"),
          "Byte Order\n排列格式(Intel/Motorola)": "",
          "Start Byte\n起始字节": "",
          "Start Bit\n起始位": "",
          "Signal Send Type\n信号发送类型": "",
          "Bit Length (Bit)\n信号长度": "",
          "Data Type\n数据类型": "",
          "Resolution\n精度": "",
          "Offset\n偏移量": "",
          "Signal Min. Value (phys)\n物理最小值": "",
          "Signal Max. Value(phys)\n物理最大值": "",
          "Signal Min. Value (Hex)\n总线最小值": "",
          "Signal Max. Value(Hex)\n总线最大值": "",
          "Initial Value(Hex)\n初始值": "",
          "Invalid Value(Hex)无效值": "",
          "Inactive Value(Hex)\n非使能值": "",
          "Unit\n单位": "",
          "Signal Value Description\n信号值描述": "",
          "Msg Cycle Time Fast(ms)\n报文发送的快速周期(ms)": this.getNewKey(msgValue, "msg_cycle_time_fast"),
          "Msg Nr. Of Repetition\n报文快速发送的次数": this.getNewKey(msgValue, "msg_nr_of_repetition"),
          "Msg Delay Time(ms)\n报文延时时间(ms)": this.getNewKey(msgValue, "msg_delay_time"),
          "源节点": this.getNewKey(msgValue, "source_node"),
          "路由": this.getNewKey(msgValue, "routing"),
          "Re_Message_ID": msgIndex,
          "Re_Signal_ID": "",
          "是否新增": msgNew
        }]
        newNewDbc = newNewDbc.concat(newMsg)
        let dbcSignalSortedMap = Array.from(this.dbcSignalMap);
        dbcSignalSortedMap.sort(function (a, b) {
          return a[0].localeCompare(b[0])
        });
        for (let [signalKey, signalValue] of dbcSignalSortedMap) {
          // dbcSignalSortedMap.forEach((signalValue, signalKey) => {
          if (signalKey.startsWith(msgKey)) {
            signalKey = this.getNewSignalKey(msgKey, signalValue);
            let signalIndex = this.baseSignalMap.get(signalKey);
            let signalNew = false;
            if (signalIndex == null) {
              signalIndex = this.dbcSignalNewIndexMap.get(signalKey)
              if (signalIndex == null) {
                console.log("signalKey", signalKey)
              } else {
                signalNew = true
              }
            }
            console.log("signalValue:", signalValue);
            console.log("signal_byte_order:", this.newDbcHeaderMap.get("signal_byte_order"));
            console.log("signal_byte_order_value:", this.getNewKey(signalValue, "signal_byte_order"));
            console.log("signal_byte_order_row:", this.getRowKey(signalValue, "排列格式"));
            const newSignal = [{
              "Signal Name\n信号名称": this.getNewKey(signalValue, "signal_name"),
              "Signal Description\n信号描述": this.getNewKey(signalValue, "signal_description"),
              "Byte Order\n排列格式(Intel/Motorola)": this.getRowKey(signalValue, "排列格式"),
              "Start Byte\n起始字节": this.getNewKey(signalValue, "signal_start_byte"),
              "Start Bit\n起始位": this.getNewKey(signalValue, "signal_start_bit"),
              "Signal Send Type\n信号发送类型": this.getNewKey(signalValue, "signal_send_type"),
              "Bit Length (Bit)\n信号长度": this.getNewKey(signalValue, "signal_bit_length"),
              "Data Type\n数据类型": this.getNewKey(signalValue, "signal_data_type"),
              "Resolution\n精度": this.getNewKey(signalValue, "signal_resolution"),
              "Offset\n偏移量": this.getNewKey(signalValue, "signal_offset"),
              "Signal Min. Value (phys)\n物理最小值": this.getNewKey(signalValue, "signal_min_value_phys"),
              "Signal Max. Value(phys)\n物理最大值": this.getNewKey(signalValue, "signal_max_value_phys"),
              "Signal Min. Value (Hex)\n总线最小值": this.getNewKey(signalValue, "signal_min_value_hex"),
              "Signal Max. Value(Hex)\n总线最大值": this.getNewKey(signalValue, "signal_max_value_hex"),
              "Initial Value(Hex)\n初始值": this.getNewKey(signalValue, "signal_initial_value_hex"),
              "Invalid Value(Hex)无效值": this.getNewKey(signalValue, "signal_invalid_value_hex"),
              "Inactive Value(Hex)\n非使能值": this.getNewKey(signalValue, "signal_inactive_value_hex"),
              "Unit\n单位": this.getNewKey(signalValue, "signal_unit"),
              "Signal Value Description\n信号值描述": this.getNewKey(signalValue, "signal_value_description"),
              "源节点": this.getNewKey(signalValue, "source_node"),
              "Re_Signal_ID": signalIndex,
              "是否新增": signalNew
            }]
            newNewDbc = newNewDbc.concat(newSignal)
          }
        }
      }
      const jsonWorkSheet = XLSX.utils.json_to_sheet(newNewDbc)
      const workBook = {
        SheetNames: ["total"],
        Sheets: {
          ["total"]: jsonWorkSheet,
        }
      };
      const today = new Date()
      const date = today.getFullYear() + "" + (today.getMonth() + 1).toString().padStart(2, "0") + "" + today.getDate().toString().padStart(2, "0")
      const fileName = "新版DBC_" + date + ".xlsx"
      XLSX.writeFile(workBook, fileName);
    },
    // 导出差异DBC
    exportDiffDbc(evt) {
      if (!this.checkBaseDbc(evt)) {
        return;
      }
      if (!this.checkNewDbc(evt)) {
        return;
      }
      for (let [msgKey, index] of this.baseMsgIndexMap) {
        if (!this.dbcMsgMatchMap.has(msgKey)) {
          this.dbcMsgDiffMap.set(msgKey, index)
        }
      }
      for (let [signalKey, index] of this.baseSignalMap) {
        if (!this.dbcSignalMatchMap.has(signalKey)) {
          this.dbcSignalDiffMap.set(signalKey, index)
        }
      }
      let newDiffDbc = [];


      let dbcSignalDiffSortedMap = Array.from(this.dbcSignalDiffMap);
      dbcSignalDiffSortedMap.sort(function (a, b) {
        return a[0].localeCompare(b[0])
      });
      let lastMsgKey = "";
      let msgKey;
      for (let [signalKey, index] of dbcSignalDiffSortedMap) {
        msgKey = signalKey.split("-")[0] + "-" + signalKey.split("-")[1] + "-" + signalKey.split("-")[2]
        if (msgKey !== lastMsgKey) {
          const newMsg = [{
            "Msg Name\n报文名称": signalKey.split("-")[0],
            "Msg ID\n报文标识符": signalKey.split("-")[1],
            "Msg Length (Byte)\n报文长度": signalKey.split("-")[2],
            "Re_Message_ID": this.baseMsgIndexMap.get(msgKey)
          }]
          newDiffDbc = newDiffDbc.concat(newMsg)
          lastMsgKey = msgKey
        }
        const newSignal = [{
          "Signal Name\n信号名称": signalKey.split("-")[3],
          "Re_Signal_ID": index
        }]
        newDiffDbc = newDiffDbc.concat(newSignal)
      }
      const jsonWorkSheet = XLSX.utils.json_to_sheet(newDiffDbc)
      const workBook = {
        SheetNames: ["total"],
        Sheets: {
          ["total"]: jsonWorkSheet,
        }
      };
      const today = new Date()
      const date = today.getFullYear() + "" + (today.getMonth() + 1).toString().padStart(2, "0") + "" + today.getDate().toString().padStart(2, "0")
      const fileName = "差异DBC_" + date + ".xlsx"
      XLSX.writeFile(workBook, fileName);
    },
    // 获取表头
    getHeaderRow(sheet) {
      const headers = [];
      const range = XLSX.utils.decode_range(sheet["!ref"]);
      let C;
      const R = range.s.r;
      for (C = range.s.c; C <= range.e.c; ++C) {
        const cell = sheet[XLSX.utils.encode_cell({c: C, r: R})];
        let hdr = "UNKNOWN " + C;
        if (cell && cell.t) hdr = XLSX.utils.format_cell(cell);
        headers.push(hdr);
      }
      return headers;
    },
    // 获取基线Key值
    getBaseKey(row, key) {
      let header = this.baseDbcHeaderMap.get(key);
      if (header == null) {
        console.error("获取key", key, "异常", row);
        return null;
      }
      let val = this.getRowKey(row, header);
      if (val != null && typeof val === 'string') {
        val = val.trim();
        val = val.replace(/[\r\n]/g, "");
      }
      return val;
    },
    // 获取新版Key值
    getNewKey(row, key) {
      let header = this.newDbcHeaderMap.get(key);
      if (header == null) {
        console.error("获取key", key, "异常", row);
        return null;
      }
      let val = this.getRowKey(row, header);
      if (val != null && typeof val === 'string') {
        val = val.trim();
        val = val.replace(/[\r\n]/g, "");
      }
      return val;
    },
    // 获取记录Key值
    getRowKey(row, key) {
      for (let rowKey in row) {
        let tmp = rowKey;
        tmp = tmp.trim();
        tmp = tmp.replace(/[\r\n]/g, "");
        tmp = tmp.replace(" ", "");
        if (tmp.indexOf(key) > -1) {
          return row[rowKey]
        }
      }
      return null;
    }
  }
}
</script>
<style lang="less">
// @import './assets/css/styles.less';
body {
  background-color: #1b1c25; /* 白色 */
  color: #afe8f7;
  padding: 10px;
}

h1 {
  padding-bottom: 5px;
  border-bottom: 1px solid #afe8f7;
}

.container {
  display: grid;
  grid-template-columns: 1fr 1fr;
}

.item {
  text-align: left;
  margin: 5px;
}

.drag-box {
  height: 100px;
  display: flex;
  justify-content: center; /* 水平居中 */
  align-items: center; /* 垂直居中 */
  border: 1px dashed #afe8f7;
  margin-bottom: 10px;
}

.function-button {
  margin-top: 50px;
  text-align: center;
}
</style>

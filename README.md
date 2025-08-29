# exportjsontoexcel
完整Excel导出工具，支持多key值合并、表头自定义对齐、空行合并

## 安装
```
npm install xlsx-js-style
npm install exportjsontoexcel
```

## 使用
```
import exportJsonToExcel from 'exportjsontoexcel';

const tableData = [
  {
    'areaCode': '320100',
    'dangerCount': '37',
    'drivingSchoolCount': '0',
    'goodsCount': '1705',
    'orgName': 'xx市江xxx区1',
    'passengerCount': '4',
    'passengerStationCount': '0',
    'statisticalItemName': '企业数',
    'summaryCount': '1708',
    'test': {
      aa: '11'
    }
  },
  {
    'areaCode': '320100',
    'dangerCount': '0',
    'drivingSchoolCount': '0',
    'goodsCount': '0',
    'orgName': 'xx市江xxx区2',
    'passengerCount': '0',
    'passengerStationCount': '0',
    'statisticalItemName': '主要负责人数',
    'summaryCount': '0',
    'test': {
      aa: '22'
    }
  },
  {
    'areaCode': '320100',
    'dangerCount': '0',
    'drivingSchoolCount': '0',
    'goodsCount': '0',
    'orgName': 'xx市江xxx区3',
    'passengerCount': '0',
    'passengerStationCount': '0',
    'statisticalItemName': '安全生产管理人员数',
    'summaryCount': '0',
    'test': {
      aa: '33'
    }
  },
  {
    'areaCode': '320100',
    'dangerCount': '0',
    'drivingSchoolCount': '0',
    'goodsCount': '0',
    'orgName': 'xx市江xxx区4',
    'passengerCount': '0',
    'passengerStationCount': '0',
    'statisticalItemName': '两类人员数',
    'summaryCount': '0',
    'test': {
      aa: '44'
    }
  }
]





exportJsonToExcel(tableData, {
  headers: [
    {
      property: '',
      title: '统计时间：2025年7月22日',
      alignment: { horizontal: 'left' },
      children: [
        { property: 'orgName', title: '管辖机构', width: 20 },
        { property: 'statisticalItemName', title: '统计项', width: 20 },
        { property: 'summaryCount', title: '汇总', width: 20 },
        { property: 'passengerCount', title: '道路旅客运输', width: 20 },
        { property: 'goodsCount', title: '道路普货运输', width: 20 },
        { property: 'dangerCount', title: '道路危险品运输', width: 20 },
        { property: 'passengerStationCount', title: '道路客运站', width: 20 },
        { property: 'drivingSchoolCount', title: '一级机动车驾驶员培训机构', width: 30 },
        { property: 'test.aa', title: '多级属性的使用', width: 30 }
      ]
    }
  ],
  // 行合并配置：自动合并分类相同的行
  rowMergeRules: [{ key: 'orgName', merge: true }],
  filename: '全省两类人员考核情况统计表',
  sheetName: '总数（省）',
  mainTitle: '全省两类人员考核情况统计表',
  // 表格底部添加数据说明
  notes: [
    '这是一条非常长的备注信息，用于测试文本换行功能。当文本长度超过表格宽度时，',
    '应该自动换行显示，而不是撑大表格列宽。这条备注会被合并成一个单元格，',
    '并且保持与表格主体相同的宽度。'
  ],
  headerStyle: {
    font: {
      sz: 12,
      name: '仿宋',
      color: { rgb: '000000' },
      bold: true
    },
    fill: {
      fgColor: { rgb: 'F2F2F2', color: { rgb: '000000' } }
    },
    alignment: {
      vertical: 'center',
      horizontal: 'center'
    }
  },
  mainTitleStyle: {
    font: {
      sz: 14,
      name: '仿宋',
      color: { rgb: '000000' },
      bold: true
    },
    alignment: {
      vertical: 'center',
      horizontal: 'center'
    }
  },
  // 完全自定义单元格样式
  getCellStyle: (rowIndex, colIndex, rowData) => {
    let style = null

    // 偶数行背景色（斑马线）
    if (rowIndex % 2 !== 0) {
      style = {
        fill: { fgColor: { rgb: 'eeeeee' } }
      }
    }

    return style
  }
})
```
 

/*
 * @description: 完整Excel导出工具，支持多key值合并、表头自定义对齐、空行合并
 * 核心功能：
 * - 表头默认居中，同时保留自定义对齐设置（如horizontal: 'left'）
 * - 支持通过property属性配置多key（格式："key1,key2,key3"），值自动换行显示
 * - 自动合并表头空行与有效表头
 * - 支持多级表头、单元格合并、自适应列宽等
 * @author: ZhuJun
 * @date: 2025-08-27
 * @version: V1.4.0
 */

// @ts-nocheck
import XLSX from 'xlsx-js-style'

/**
 * 导出JSON数据到Excel文件
 * @param {Array} jsonData - 要导出的JSON数据（非空数组）
 * @param {Object} options - 导出配置选项
 * @returns {void}
 */
function exportJsonToExcel(jsonData, options = {}) {
  // 基础数据验证
  if (!Array.isArray(jsonData) || jsonData.length === 0) {
    throw new Error('导出数据不能为空，请提供有效的JSON数组')
  }

  // 解构配置选项，设置默认值
  const {
    // 备注相关配置
    notes = [],
    notesStyle = createDefaultStyle({
      font: { name: '微软雅黑', sz: 10, color: { rgb: 'FF666666' } },
      alignment: { horizontal: 'left', vertical: 'top', wrapText: true },
      border: undefined,
      fill: { fgColor: { rgb: 'FFFFFF' } },
    }),
    notesMerge = true,
    notesSeparator = '\n',

    // 合并规则配置
    rowMergeRules = [],
    mergeCells = [],

    // 表头与文件名配置
    headers = null,
    filename = '导出数据',
    sheetName = 'Sheet1',
    mainTitle = '',

    // 样式配置
    mainTitleStyle = createDefaultStyle({
      font: { name: '宋体', sz: 16, color: { rgb: 'FF000080' }, bold: true },
      alignment: { horizontal: 'center', vertical: 'center' },
      fill: { fgColor: { rgb: 'FFFFFF' } },
    }),
    headerStyle = createDefaultStyle({
      font: { name: '宋体', sz: 13, color: { rgb: 'FF000000' }, bold: true },
      alignment: { horizontal: 'center', vertical: 'center' }, // 默认居中
      fill: { fgColor: { rgb: '90C3EA' } },
    }),
    cellStyle = createDefaultStyle({
      font: { name: '宋体', sz: 12, color: { rgb: 'FF333333' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      fill: { fgColor: { rgb: 'FFFFFF' } },
    }),
    borderStyle = {
      border: {
        top: { style: 'thin', color: { rgb: 'FF000000' } },
        bottom: { style: 'thin', color: { rgb: 'FF000000' } },
        left: { style: 'thin', color: { rgb: 'FF000000' } },
        right: { style: 'thin', color: { rgb: 'FF000000' } },
      },
    },
    getCellStyle = null,

    // 列宽配置
    autoWidth = true,
    widthMultiplier = 1.1,
  } = options

  // 初始化工作簿和工作表相关变量
  const wb = XLSX.utils.book_new()
  let ws, wsData, allMerges = []

  // 处理备注数据
  const rawNotes = Array.isArray(notes) ? notes : [notes]
  const validNotes = rawNotes.filter(note => note != null && note !== '')
  const hasNotes = validNotes.length > 0

  // 处理表头和数据
  let headerRows = [], dataRows = [], columnCount = 0,
    leafHeaders = [], headerRowsLength = 0,
    columnAlignments = [], mergeAlignments = {}

  if (headers && headers.length > 0) {
    // 处理多级表头 - 保留自定义对齐
    const processedHeaders = processMultiLevelHeaders(headers, headerStyle)
    headerRows = processedHeaders.headerRows
    leafHeaders = processedHeaders.leafHeaders
    columnCount = processedHeaders.columnCount
    columnAlignments = processedHeaders.columnAlignments

    // 生成数据行 - 支持多key值合并（从property获取多个key）
    dataRows = jsonData.map(row =>
      leafHeaders.map(header => {
        // 处理多key值合并（property属性，格式："key1,key2"）
        if (header.property) {
          const keys = header.property.split(',').map(k => k.trim())
          // 获取每个key对应的值，过滤空值，用换行符连接
          const values = keys
            .map(key => getNestedValue(row, key))
            .filter(value => value !== '' && value != null)
          return values.join('\n')
        }
        // 兼容原key方式
        return header.key ? getNestedValue(row, header.key) : ''
      })
    )

    headerRowsLength = headerRows.length
  } else {
    // 处理单级表头（自动生成表头）
    const headerRow = Object.keys(jsonData[0]).map(key => ({
      key, title: key,
      alignment: headerStyle.alignment // 使用默认对齐
    }))
    leafHeaders = headerRow
    columnAlignments = headerRow.map(h => h.alignment)
    dataRows = jsonData.map(row => headerRow.map(h => row[h.key]))
    columnCount = headerRow.length
    headerRows = [headerRow.map(h => h.title)]
    headerRowsLength = 1
  }

  // 构建备注数据结构
  let noteData = [], noteRowCount = 0, noteContent = ''
  if (hasNotes) {
    noteContent = validNotes.join(notesSeparator)
    noteData = [new Array(columnCount).fill('')]
    noteData[0][0] = noteContent
    noteRowCount = noteData.length
  }

  // 构建完整工作表数据
  wsData = [[mainTitle], ...headerRows, ...dataRows]
  if (hasNotes) wsData = [...wsData, [], ...noteData]
  ws = XLSX.utils.aoa_to_sheet(wsData)

  // 主标题合并（横跨所有列）
  const mainTitleMerge = {
    s: { r: 0, c: 0 },
    e: { r: 0, c: columnCount - 1 }
  }
  applyStyleToCell(ws, { r: 0, c: 0 }, mainTitleStyle)

  // 生成表头合并规则
  let headerMerges = []
  if (headers && headers.length > 0) {
    const headerMergeResult = generateHeaderMerges(headers, headerStyle)
    headerMerges = headerMergeResult.merges
    mergeAlignments = headerMergeResult.alignments
  }

  // 生成表头空行合并规则（纵向合并）
  const emptyRowMerges = headers && headers.length > 0
    ? detectEmptyRowMerges(headerRows, columnCount, 1, headerRowsLength)
    : []

  // 生成数据行合并规则
  const rowMerges = generateRowMerges({
    dataRows, jsonData, rowMergeRules, leafHeaders, headerRowsLength
  })

  // 生成备注区域合并规则
  const noteMerges = []
  if (hasNotes && notesMerge && columnCount > 0) {
    const noteStartRow = 1 + headerRowsLength + dataRows.length + 1
    noteMerges.push({
      s: { r: noteStartRow, c: 0 },
      e: { r: noteStartRow + noteRowCount - 1, c: columnCount - 1 }
    })
  }

  // 合并所有合并规则并过滤无效合并
  allMerges = [
    mainTitleMerge,
    ...headerMerges,
    ...emptyRowMerges,
    ...mergeCells,
    ...rowMerges,
    ...noteMerges
  ]
  ws['!merges'] = allMerges.filter(merge =>
    (merge.s.r !== merge.e.r || merge.s.c !== merge.e.c) &&
    merge.s.r <= merge.e.r && merge.s.c <= merge.e.c
  )

  // 应用表头样式 - 保留自定义对齐
  for (let r = 1; r <= headerRowsLength; r++) {
    for (let c = 0; c < columnCount; c++) {
      const alignment = columnAlignments[c] || headerStyle.alignment
      const currentHeaderStyle = {
        ...headerStyle,
        ...borderStyle,
        alignment: alignment
      }
      applyStyleToCell(ws, { r, c }, currentHeaderStyle)
    }
  }

  // 应用数据行样式 - 确保换行生效
  const dataStartRow = headerRowsLength + 1
  const dataRange = XLSX.utils.decode_range(ws['!ref'])
  for (let r = dataStartRow; r <= dataRange.e.r; r++) {
    const rowIndex = r - dataStartRow
    if (rowIndex >= dataRows.length) break
    for (let c = 0; c <= dataRange.e.c; c++) {
      // 确保wrapText为true，使换行符生效
      const baseStyle = {
        ...cellStyle,
        ...borderStyle,
        alignment: { ...cellStyle.alignment, wrapText: true }
      }
      const customStyle = getCellStyle?.(rowIndex, c, jsonData[rowIndex])
      applyStyleToCell(ws, { r, c }, customStyle ? { ...baseStyle, ...customStyle } : baseStyle)
    }
  }

  // 处理合并单元格对齐
  if (ws['!merges']?.length) {
    ws['!merges'].forEach(merge => {
      const mergeKey = `${merge.s.r}-${merge.s.c}`
      if (mergeAlignments[mergeKey]) {
        applyStyleToCell(ws, { r: merge.s.r, c: merge.s.c }, {
          alignment: { ...mergeAlignments[mergeKey], vertical: 'center', wrapText: true }
        })
      } else if (merge.s.r >= 1 && merge.s.r <= headerRowsLength) {
        const alignment = columnAlignments[merge.s.c] || headerStyle.alignment
        applyStyleToCell(ws, { r: merge.s.r, c: merge.s.c }, {
          alignment: { ...alignment, vertical: 'center', wrapText: true }
        })
      } else {
        applyStyleToCell(ws, { r: merge.s.r, c: merge.s.c }, {
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
        })
      }
    })
  }

  // 应用备注区域样式
  if (hasNotes) {
    const noteStartRow = 1 + headerRowsLength + dataRows.length + 1
    applyStyleToCell(ws, { r: noteStartRow, c: 0 }, {
      ...notesStyle,
      border: undefined
    })
  }

  // 处理列宽 - 多值情况下需要更宽的列宽
  if (autoWidth || headers?.some(h => h.width !== undefined)) {
    const dataOnlyRange = {
      s: { r: 0, c: 0 },
      e: { r: 1 + headerRowsLength + dataRows.length - 1, c: columnCount - 1 }
    }
    const calculatedWidths = autoWidth ? calculateColumnWidths(ws, headers, dataOnlyRange) : []
    const finalWidths = []
    for (let i = 0; i < columnCount; i++) {
      const leafHeader = headers?.some(h => h.children) ? getLeafHeaders(headers)[i] : headers?.[i]
      // 多key列默认增加宽度系数
      const widthFactor = leafHeader?.property ? 1.5 : widthMultiplier
      finalWidths[i] = leafHeader?.width ?? (autoWidth ? calculatedWidths[i] * widthFactor : 10)
    }
    ws['!cols'] = finalWidths.map(width => ({ wch: Math.ceil(width) }))
  }

  // 处理备注行高度自适应
  if (hasNotes) {
    const noteStartRow = 1 + headerRowsLength + dataRows.length + 1
    const lineHeight = 15 // 每行高度（像素）
    const lineCount = Math.max(1, (noteContent.match(/\n/g) || []).length + 1)
    const contentWidth = ws['!cols']?.reduce((sum, col) => sum + col.wch, 0) || 100
    const charPerLine = Math.floor(contentWidth / 2) // 估算每行可容纳字符数
    let actualLineCount = 0
    noteContent.split('\n').forEach(line => {
      actualLineCount += Math.max(1, Math.ceil(line.length / charPerLine))
    })
    if (!ws['!rows']) ws['!rows'] = []
    ws['!rows'][noteStartRow] = { hpx: actualLineCount * lineHeight }
  }

  // 确保工作表范围有效，避免内容截断
  ensureValidRange(ws, dataRows.length, headerRowsLength, hasNotes, noteRowCount, columnCount)

  // 导出文件
  XLSX.utils.book_append_sheet(wb, ws, sheetName)
  XLSX.writeFile(wb, `${filename}.xlsx`)
}

/**
 * 处理多级表头，支持property属性（多key配置）和自定义对齐
 */
function processMultiLevelHeaders(headers, headerStyle) {
  const headerRows = []
  const leafHeaders = []
  const columnAlignments = []
  let maxLevel = 0

  function collectLeafHeaders(headerList, parentPath = [], level = 1) {
    if (level > maxLevel) maxLevel = level

    headerList.forEach(header => {
      // 合并默认对齐与自定义对齐（自定义属性优先）
      const headerAlignment = header.alignment
        ? { ...headerStyle.alignment, ...header.alignment }
        : headerStyle.alignment;

      // 保留property信息（多key配置）和key信息
      const currentHeader = {
        ...header,
        alignment: headerAlignment,
        property: header.property?.trim(), // 存储多key配置
        key: header.key?.trim() // 保留单key配置
      }

      const currentPath = [...parentPath, currentHeader]

      if (currentHeader.children && currentHeader.children.length > 0) {
        collectLeafHeaders(currentHeader.children, currentPath, level + 1)
      } else {
        leafHeaders.push({ ...currentHeader, path: currentPath })
        columnAlignments.push(currentHeader.alignment)
      }
    })
  }

  collectLeafHeaders(headers)

  // 初始化表头行
  for (let i = 0; i < maxLevel; i++) {
    headerRows.push(Array(leafHeaders.length).fill(''))
  }

  // 填充表头行数据
  leafHeaders.forEach((leafHeader, leafIndex) => {
    leafHeader.path.forEach((header, levelIndex) => {
      const rowIndex = levelIndex
      if (headerRows[rowIndex][leafIndex] === '') {
        headerRows[rowIndex][leafIndex] = header.title
      }
    })
  })

  return {
    headerRows,
    leafHeaders,
    columnCount: leafHeaders.length,
    columnAlignments
  }
}

/**
 * 生成表头横向合并规则
 */
function generateHeaderMerges(headers, headerStyle) {
  const merges = []
  const mergeAlignments = {}

  function processLevel(headerList, level, startCol) {
    let currentCol = startCol

    headerList.forEach(header => {
      // 合并默认对齐与自定义对齐
      const alignment = header.alignment
        ? { ...headerStyle.alignment, ...header.alignment }
        : headerStyle.alignment;

      const span = header.children ? calculateColumnCount(header.children) : 1

      if (span > 1) {
        const mergeKey = `${level}-${currentCol}`
        mergeAlignments[mergeKey] = alignment

        merges.push({
          s: { r: level, c: currentCol },
          e: { r: level, c: currentCol + span - 1 }
        })
      }

      if (header.children) {
        processLevel(header.children, level + 1, currentCol)
      }

      currentCol += span
    })
  }

  processLevel(headers, 1, 0)
  return { merges, alignments: mergeAlignments }
}

/**
 * 检测表头空行并与有效表头合并（纵向合并）
 */
function detectEmptyRowMerges(headerRows, columnCount, headerStartRow, headerTotalRows) {
  const merges = []

  for (let col = 0; col < columnCount; col++) {
    const columnCells = headerRows.map((row, rowIdx) => ({
      value: row[col],
      rowIdx
    }))

    const nonEmptyCells = columnCells.filter(cell =>
      cell.value !== '' && cell.value != null
    )

    if (nonEmptyCells.length === 0) continue

    // 倒序处理有效表头，确保底层优先合并
    for (let i = nonEmptyCells.length - 1; i >= 0; i--) {
      const currentCell = nonEmptyCells[i]
      const currentRowIdx = currentCell.rowIdx

      // 向上查找连续空行
      let startMergeRowIdx = currentRowIdx
      for (let r = currentRowIdx - 1; r >= 0; r--) {
        const cell = columnCells[r]
        if (cell.value === '' || cell.value == null) {
          startMergeRowIdx = r
        } else {
          break
        }
      }

      // 向下查找连续空行
      let endMergeRowIdx = currentRowIdx
      for (let r = currentRowIdx + 1; r < headerTotalRows; r++) {
        const cell = columnCells[r]
        if (cell.value === '' || cell.value == null) {
          endMergeRowIdx = r
        } else {
          break
        }
      }

      // 转换为Excel行索引
      const startExcelRow = headerStartRow + startMergeRowIdx
      const endExcelRow = headerStartRow + endMergeRowIdx

      if (startExcelRow < endExcelRow) {
        merges.push({
          s: { r: startExcelRow, c: col },
          e: { r: endExcelRow, c: col }
        })
      }
    }
  }

  return merges
}

/**
 * 确保工作表范围有效，避免内容被截断
 */
function ensureValidRange(ws, dataRowsLength, headerRowsLength, hasNotes, noteRowCount, columnCount) {
  const maxRowIndex = 1 + headerRowsLength + dataRowsLength + (hasNotes ? 1 + noteRowCount : 0)
  const maxColIndex = columnCount - 1

  const currentRange = ws['!ref']
    ? XLSX.utils.decode_range(ws['!ref'])
    : { s: { r: 0, c: 0 }, e: { r: maxRowIndex, c: maxColIndex } }

  const newRange = {
    s: { r: 0, c: 0 },
    e: {
      r: Math.max(currentRange.e.r, maxRowIndex),
      c: Math.max(currentRange.e.c, maxColIndex)
    }
  }

  ws['!ref'] = XLSX.utils.encode_range(newRange)
}

/**
 * 计算列宽（考虑换行内容，取最长行宽度）
 */
function calculateColumnWidths(ws, headers, dataRange) {
  const columnCount = dataRange.e.c - dataRange.s.c + 1
  const columnWidths = Array(columnCount).fill(0)

  for (let r = dataRange.s.r; r <= dataRange.e.r; r++) {
    for (let c = dataRange.s.c; c <= dataRange.e.c; c++) {
      const cellAddress = XLSX.utils.encode_cell({ r, c })
      const cell = ws[cellAddress]

      if (cell && cell.v !== undefined) {
        const cellValue = String(cell.v)
        // 对于包含换行的单元格，取最长行的宽度
        const lines = cellValue.split('\n')
        const maxLineWidth = lines.reduce((max, line) => {
          const lineWidth = getStringWidth(line)
          return Math.max(max, lineWidth)
        }, 0)
        if (maxLineWidth > columnWidths[c]) {
          columnWidths[c] = maxLineWidth
        }
      }
    }
  }

  return columnWidths
}

/**
 * 生成数据行合并规则
 */
function generateRowMerges({ dataRows, jsonData, rowMergeRules, leafHeaders, headerRowsLength }) {
  const merges = []
  const dataStartRow = 1 + headerRowsLength

  // 函数式合并规则
  if (typeof rowMergeRules === 'function') {
    for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
      for (let colIndex = 0; colIndex < dataRows[rowIndex].length; colIndex++) {
        const mergeCount = rowMergeRules(rowIndex, colIndex, jsonData[rowIndex])
        if (mergeCount > 1 && rowIndex + mergeCount <= dataRows.length) {
          merges.push({
            s: { r: dataStartRow + rowIndex, c: colIndex },
            e: { r: dataStartRow + rowIndex + mergeCount - 1, c: colIndex }
          })
        }
      }
    }
  }
  // 数组式合并规则（按字段合并相同值）
  else if (Array.isArray(rowMergeRules)) {
    rowMergeRules.forEach(rule => {
      if (!rule.key || !rule.merge) return

      const colIndex = leafHeaders.findIndex(header =>
        header.key === rule.key || (header.property && header.property.includes(rule.key))
      )
      if (colIndex === -1) return

      let startRow = 0
      // 对于多key列，使用第一个key的值进行合并判断
      const targetKey = rule.key || leafHeaders[colIndex].property?.split(',')[0]
      let currentValue = targetKey ? getNestedValue(jsonData[0], targetKey) : ''

      for (let rowIndex = 1; rowIndex < jsonData.length; rowIndex++) {
        const value = targetKey ? getNestedValue(jsonData[rowIndex], targetKey) : ''

        if (value !== currentValue || rowIndex === jsonData.length - 1) {
          const endRow = rowIndex === jsonData.length - 1 && value === currentValue
            ? rowIndex
            : rowIndex - 1

          if (endRow - startRow >= 1) {
            merges.push({
              s: { r: dataStartRow + startRow, c: colIndex },
              e: { r: dataStartRow + endRow, c: colIndex }
            })
          }

          startRow = rowIndex
          currentValue = value
        }
      }
    })
  }

  return merges
}

/**
 * 获取嵌套对象的属性值（支持"a.b.c"形式的键）
 */
function getNestedValue(obj, key) {
  return key.split('.').reduce((o, i) => (o && o[i] !== undefined ? o[i] : ''), obj)
}

/**
 * 计算表头包含的总列数
 */
function calculateColumnCount(headers) {
  return headers.reduce((acc, header) => {
    return acc + (header.children ? calculateColumnCount(header.children) : 1)
  }, 0)
}

/**
 * 获取所有叶子表头（最底层表头）
 */
function getLeafHeaders(headers) {
  const result = []

  function collect(headersList) {
    headersList.forEach(header => {
      if (header.children && header.children.length > 0) {
        collect(header.children)
      } else {
        result.push(header)
      }
    })
  }

  collect(headers)
  return result
}

/**
 * 应用样式到指定单元格
 */
function applyStyleToCell(worksheet, { r, c }, style) {
  const cellAddress = XLSX.utils.encode_cell({ r, c })
  const cell = worksheet[cellAddress] || {}

  worksheet[cellAddress] = {
    ...cell,
    s: { ...(cell.s || {}), ...style }, // 合并已有样式和新样式
    v: cell.v !== undefined ? cell.v : '' // 确保单元格有值
  }
}

/**
 * 计算字符串宽度（考虑中文字符，中文字符按2个字符宽度计算）
 */
function getStringWidth(str) {
  if (!str) return 0
  return str.replace(/[^\x00-\xff]/g, '**').length + 2 // +2 预留边距
}

/**
 * 创建默认样式（可通过customStyle覆盖）
 */
function createDefaultStyle(customStyle = {}) {
  return {
    font: {
      name: '微软雅黑',
      sz: 11,
      color: { rgb: 'FF000000' },
      ...(customStyle.font || {})
    },
    alignment: {
      horizontal: 'left',
      vertical: 'top',
      wrapText: false,
      ...(customStyle.alignment || {})
    },
    fill: { ...(customStyle.fill || {}) },
    border: {
      top: { style: 'thin' },
      bottom: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' },
      ...(customStyle.border || {})
    },
    ...customStyle
  }
}


export default exportJsonToExcel

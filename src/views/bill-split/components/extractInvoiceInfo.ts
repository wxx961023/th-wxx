const extractInvoiceInfo = (text: string, pageNum: number) => {
  const extractedData: any[] = []

  console.log(`正在解析第${pageNum}页文本，长度:`, text.length)

  // 提取印刷序号
  const invoiceNumbers = new Set<string>()

  // 特殊处理：直接提取最长的数字字母混合字符串，然后按位置截取
  console.log('🎯 开始尝试最长字符串截取方法...')

  // 1. 先清理文本，移除空格
  const cleanedText = text.replace(/\s+/g, '')
  console.log('清理后的文本（无空格）:', cleanedText)

  // 2. 查找最长的数字字母混合字符串（至少25位）
  const mixedStringPattern = /[A-Z0-9]{25,}/g
  const mixedStringMatches = cleanedText.match(mixedStringPattern)

  if (mixedStringMatches && mixedStringMatches.length > 0) {
    console.log('找到数字字母混合字符串:', mixedStringMatches)

    // 找到最长的字符串
    const longestString = mixedStringMatches.reduce((a, b) => a.length > b.length ? a : b)
    console.log(`🎉 找到最长的字符串: "${longestString}" (长度: ${longestString.length})`)

    // 3. 简化逻辑：只使用两个有效的方法
    // 1. 位置截取：从第3位截取13位作为电子客票号码
    const ticketNumber = longestString.substring(2, 15)
    invoiceNumbers.add(ticketNumber)
    console.log(`✅ 位置截取电子客票号码: "${ticketNumber}" (第3-15位)`)

    // 2. 模式匹配：在整个文本中查找20位数字作为发票号码
    const allInvoiceMatches = cleanedText.match(/\d{20}/g)
    if (allInvoiceMatches) {
      allInvoiceMatches.forEach(invoice => {
        invoiceNumbers.add(invoice)
        console.log(`✅ 模式匹配发票号码: "${invoice}"`)
      })
    }
  } else {
    console.log('❌ 未找到足够长的数字字母混合字符串')
  }

  console.log(`印刷序号提取结果: ${Array.from(invoiceNumbers).length}个`)
  console.log('提取到的印刷序号:', Array.from(invoiceNumbers))

  // 组合数据 - 区分电子客票号和发票号码
  const invoiceArray = Array.from(invoiceNumbers)

  // 合并数据为单条记录，用不同字段存储
  if (invoiceArray.length > 0) {
    let ticketNumber = null; // 13位电子客票号
    let invoiceNumber = null; // 20位发票号码

    // 遍历提取的数据，分类存储
    invoiceArray.forEach((invoice, index) => {
      console.log(`处理发票数据 ${index + 1}: "${invoice}" (长度: ${invoice.length})`)

      if (invoice.length === 13) {
        ticketNumber = invoice;
        console.log(`  ✅ 识别为电子客票号: "${ticketNumber}"`);
      } else if (invoice.length === 20) {
        invoiceNumber = invoice;
        console.log(`  ✅ 识别为发票号码: "${invoiceNumber}"`);
      }
    });

    // 创建合并后的单条记录
    extractedData.push({
      ticketNumber: ticketNumber, // 13位电子客票号
      invoiceNumber: invoiceNumber, // 20位发票号码
      originalValue: ticketNumber || invoiceNumber, // 优先使用电子客票号用于匹配
      remark: '', // 暂时不提取备注
      confidence: 1.0, // 直接提取给最高置信度
      pageNum: pageNum
    });

    console.log(`📝 合并后的记录:`, {
      ticketNumber: ticketNumber,
      invoiceNumber: invoiceNumber,
      originalValue: ticketNumber || invoiceNumber
    });
  }

  console.log(`=== 第${pageNum}页提取总结 ===`)
  console.log(`印刷序号数量: ${invoiceArray.length}`)
  console.log(`最终提取记录数: ${extractedData.length}`)

  if (extractedData.length > 0) {
    console.log('提取的详细数据:')
    extractedData.forEach((data, index) => {
      console.log(`  记录 ${index + 1}:`, {
        ticketNumber: data.ticketNumber,
        invoiceNumber: data.invoiceNumber,
        remark: data.remark,
        confidence: data.confidence,
        pageNum: data.pageNum
      })
    })
  } else {
    console.log('❌ 未提取到任何有效数据')
    console.log('💡 建议：检查PDF文本中是否包含发票号码或关键词')
  }

  return extractedData
}

export default extractInvoiceInfo

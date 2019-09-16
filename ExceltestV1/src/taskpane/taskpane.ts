/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    // document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;
    document.getElementById("setup").onclick = setUp;
    document.getElementById("add-btn").onclick = testBtn;
    // 借方减贷方
    document.getElementById("B-L").onclick = borrowersToLenders;
    // 贷方减借方
    document.getElementById("L-B").onclick = lendersToBorrowers;

    //// 所者者权益变动表
    document.getElementById("equitySheet").onclick =setEquitySheet ;
    //// 负债资产表
    document.getElementById("balanceSheet").onclick = setBalanceSheet;
    //// 利润表
    document.getElementById("profitSheet").onclick = setProfitSheet;
    //// 现金流量表
    document.getElementById("cashFlowSheet").onclick = setCashFlowSheet;
    //// 能力指标表
    document.getElementById("analyseSheet").onclick = setAnalysisSheet;
    // $("#run").click(() => tryCatch(run));
  }
});

// export async function run() {
//   try {
//     await Excel.run(async context => {
//       /**
//        * Insert your Excel code here
//        */
//       const range = context.workbook.getSelectedRange();

//       // Read the range address
//       range.load("address");

//       // Update the fill color
//       range.format.fill.color = "yellow";

//       await context.sync();
//       console.log(`The range address was ${range.address}.`);
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }

///调试
async function testBtn() {
  await Excel.run(async (context) => {
    const mainSheetName = "会计分录和分类账";
    var mainSheet = context.workbook.worksheets.getItem(mainSheetName);

    const activeRange = context.workbook.getActiveCell();
    var shapes = mainSheet.shapes;

    await context.sync().then(function() {});
  });
}

/////-----------能力指标分析表--------------------////
async function setAnalysisSheet() {
  await Excel.run(async (context) => {
    const sheetName = "能力指标分析";
    context.workbook.worksheets.getItemOrNullObject(sheetName).delete();
    const sheet = context.workbook.worksheets.add(sheetName);

    ////// 利润表
    const profitSheetName = "利润表";
    var profitSheet = context.workbook.worksheets.getItem(profitSheetName);
    ////// 资产负债表
    const balanceSheetName = "资产负债表";
    var balanceSheet = context.workbook.worksheets.getItem(balanceSheetName);

    ////第一行
    sheet.getRange("A1:B1").values = [["编制单位：", "甲公司"]];
    sheet.getRange("B1:C1").merge(); ////合并单元格

    ////表格标题
    var titleRowRange = sheet.getRange("A2:C2");
    var titleData = [["", "指标及说明", "=利润表!B3"]]; //2017为利润表的时间
    titleRowRange.values = titleData;
    titleRowRange.format.fill.color = "#79A25E";
    titleRowRange.format.font.color = "white";

    ////偿债能力
    var solvencyRange = sheet.getRange("A3:C8");
    var solvencyData = [
      ["偿债能力", "流动比率＝流动资产÷流动负债×100％", ""],
      ["", "速动比率＝速动资产÷流动负债×100％", ""],
      ["", "现金比率 = 货币资金÷流动负债×100％", ""],
      ["", " 资产负债率＝负债总额÷资产总额×100％", ""],
      ["", "产权比率＝负债总额÷所有者权益总额×100％", ""],
      ["", "权益乘数 = 资产总额÷所有者权益", ""]
    ];
    solvencyRange.values = solvencyData;

    /////营运能力
    var operationalRange = sheet.getRange("A10:C15");
    var operationalData = [
      ["营运能力", "存货周转率＝销售成本÷平均存货", ""],
      ["", "应收账款周转率＝赊销收入÷应收账款和应收票据的平均余额", ""],
      ["", "流动资产周转率＝销售收入÷平均流动资产", ""],
      ["", " 固定资产周转率＝销售收入净额÷固定资产平均净值", ""],
      ["", "总资产周转率＝销售收入净额÷平均资产余额", ""],
      ["", "应付账款周转率＝营业成本÷应付账款和应付票据平均余额", ""]
    ];
    operationalRange.values = operationalData;

    ////盈利能力
    var profitabilityRange = sheet.getRange("A17:C24");
    var profitabilityData = [
      ["盈利能力", "营业毛利率＝（营业收入－营业成本）÷营业收入×100％", ""],
      ["", "核心利润率 =（营业收入－营业成本 - 税金及附加 - 三项费用）÷营业收入×100％", ""],
      ["", "营业利润率＝营业利润÷营业收入×100％", ""],
      ["", " 营业净利率＝本期净利润÷营业收入×100％", ""],
      ["", "总资产报酬率(ROA) ＝息税前利润总额÷平均资产额×100％", ""],
      ["", "资产净利率＝净利润÷平均资产额×100％", ""],
      ["", "成本费用利润率＝利润总额÷成本费用总额", ""],
      ["", "净资产利润率（ROE) ＝净利润÷平均净资产", ""]
    ];
    profitabilityRange.values = profitabilityData;

    sheet.getRange("C3").formulas = [["=(资产负债表!B20)/(资产负债表!E20)"]];
    sheet.getRange("C4").formulas = [["=(资产负债表!B20-资产负债表!B15)/资产负债表!E20"]];
    sheet.getRange("C5").formulas = [["=资产负债表!B6/资产负债表!E20"]];
    sheet.getRange("C6").formulas = [["=资产负债表!E33/资产负债表!B45"]];
    sheet.getRange("C7").formulas = [["=资产负债表!E33/资产负债表!E44"]];
    sheet.getRange("C8").formulas = [["=资产负债表!B45/资产负债表!E44"]];
    sheet.getRange("C10").formulas = [["=利润表!B6/(资产负债表!B15)"]];
    sheet.getRange("C11").formulas = [["=利润表!B5/(资产负债表!B10)"]];
    sheet.getRange("C12").formulas = [["=利润表!B5/(资产负债表!B20)"]];
    sheet.getRange("C13").formulas = [["=利润表!B5/(资产负债表!B27)"]];

    sheet.getRange("C14").formulas = [["=利润表!B5/资产负债表!B27"]];
    sheet.getRange("C15").formulas = [["=利润表!B5/(资产负债表!B45)"]];
    sheet.getRange("C17").formulas = [["=1-利润表!B6/利润表!B5"]];
    sheet.getRange("C18").formulas = [["=(利润表!B5-利润表!B6-利润表!B8-利润表!B9-利润表!B10-利润表!B11)/利润表!B5"]];
    sheet.getRange("C19").formulas = [["=利润表!B19/利润表!B5"]];
    sheet.getRange("C20").formulas = [["=利润表!C26"]];
    sheet.getRange("C21").formulas = [["=(利润表!B24+利润表!B11)/资产负债表!B45"]];
    sheet.getRange("C22").formulas = [["=利润表!B26/(资产负债表!B45)"]];
    sheet.getRange("C23").formulas = [["=利润表!B24/(利润表!B6+利润表!B8+利润表!B9+利润表!B10+利润表!B11)"]];
    sheet.getRange("C24").formulas = [["=利润表!B26/(资产负债表!E44)"]];

    sheet.getRange("C:C").numberFormat = [["0.00%"]];

    // 合并单元格
    sheet.getRange("A3:A8").merge();
    sheet.getRange("A10:A15").merge();
    sheet.getRange("A17:A24").merge();

    // 自动换行
    sheet.getRange("B:B").format.columnWidth = 350;
    sheet.getRange("B:B").format.wrapText = true;

    ////修改样式
    sheet.getRange("A1").getEntireRow().format.rowHeight = 24;
    sheet.getRanges("A9:C9, A16:C16, A25:C25").format.fill.color = "#c5e0b4";
    //字体
    var fontRange = sheet.getRanges("B2:C2, A3:A24, A1:C1").format;
    fontRange.font.bold = true;
    fontRange.verticalAlignment = "Center";
    fontRange.horizontalAlignment = "Center";

    //设置边框
    setRangeBorder(solvencyRange);
    setRangeBorder(operationalRange);
    setRangeBorder(profitabilityRange);

    ////自动调整行高
    sheet.getUsedRange().format.font.size = 12;
    sheet.getUsedRange().format.font.name = "宋体";
    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();
    sheet.getRange("A1:C2").format.rowHeight = 25;
    sheet.getRange("C1").format.columnWidth = 50;

    await context.sync();
    sheet.activate();
  });
}
/////设置边框
async function setRangeBorder(range) {
  await Excel.run(async (context) => {
    var rangeBorder = range.format.borders;
    rangeBorder.getItem("EdgeTop").color = "black";
    rangeBorder.getItem("EdgeTop").style = "Continuous";
    rangeBorder.getItem("EdgeTop").weight = "Medium";
    rangeBorder.getItem("EdgeBottom").color = "black";
    rangeBorder.getItem("EdgeBottom").style = "Continuous";
    rangeBorder.getItem("EdgeBottom").weight = "Medium";
    rangeBorder.getItem("EdgeRight").color = "black";
    rangeBorder.getItem("EdgeRight").style = "Continuous";
    rangeBorder.getItem("EdgeRight").weight = "Medium";

    await context.sync();
  });
}
/////-----------能力指标分析表--------------------////

//////-------新建现金流量表------//////
async function setCashFlowSheet() {
  await Excel.run(async (context) => {
    ///////------新建表，若表存在，则删除
    const cashFlowSheetName = "现金流量表";
    context.workbook.worksheets.getItemOrNullObject(cashFlowSheetName).delete();
    const sheet = context.workbook.worksheets.add(cashFlowSheetName);

    ////// 利润表
    const profitSheetName = "利润表";
    var profitSheet = context.workbook.worksheets.getItem(profitSheetName);
    ////// 资产负债表
    const balanceSheetName = "资产负债表";
    var balanceSheet = context.workbook.worksheets.getItem(balanceSheetName);

    //第一行
    var firstCell = sheet.getRange("A1");
    firstCell.values = [["现金流量表"]];
    firstCell.format.font.bold = true;
    sheet.getRange("B1").values = [["企业03表"]];

    // 第二行
    var secondCell = sheet.getRange("A2:B2");
    secondCell.values = [["编制单位", "XX公司"]];

    // 第三行/四   标题行
    var thirdRow = sheet.getRange("A3:B3");
    thirdRow.values = [["项目", "=资产负债表!B3"]];
    thirdRow.format.font.bold = true;

    ///主体内容
    ////第一部分
    var partIRange = sheet.getRange("A5:A15");
    var partIValues = [
      [" 一、经营活动产生的现金流量："],
      ["    销售商品、提供劳务收到的现金"],
      ["    收到的税费返还"],
      ["    收到其他与经营活动有关的现金"],
      ["    经营活动现金流入小计"],
      ["    购买商品、接受劳务支付的现金"],
      ["    支付给职工以及为职工支付的现金"],
      ["    支付的各项税费"],
      ["    支付其他与经营活动有关的现金"],
      ["经营活动现金流出小计"],
      ["经营活动产生的现金流量净额"]
    ];
    partIRange.values = partIValues;

    ////第二部分
    var partIIRange = sheet.getRange("A16:A28");
    var partIIValues = [
      [" 二、投资活动产生的现金流量："],
      ["     收回投资收到的现金"],
      ["     取得投资收益收到的现金"],
      ["     处置固定资产、无形资产和其他长期资产收回的现金净额"],
      ["     处置子公司及其他营业单位收到的现金净额"],
      ["     收到其他与投资活动有关的现金"],
      [" 投资活动现金流入小计"],
      ["    购建固定资产、无形资产和其他长期资产支付的现金"],
      ["    投资支付的现金"],
      ["    取得子公司及其他营业单位支付的现金净额"],
      ["    支付其他与投资活动有关的现金"],
      [" 投资活动现金流出小计"],
      [" 投资活动产生的现金流量净额"]
    ];
    partIIRange.values = partIIValues;

    ////第三部分
    var partIIIRange = sheet.getRange("A29:A38");
    var partIIIValues = [
      [" 三、筹资活动产生的现金流量："],
      ["     吸收投资收到的现金"],
      ["     取得借款收到的现金"],
      ["     收到其他与筹资活动有关的现金"],
      [" 筹资活动现金流入小计"],
      ["     偿还债务支付的现金"],
      ["     分配股利、利润或偿付利息支付的现金"],
      ["     支付其他与筹资活动有关的现金"],
      [" 筹资活动现金流出小计"],
      [" 筹资活动产生的现金流量净额"]
    ];
    partIIIRange.values = partIIIValues;

    ////第四部分
    var partIVRange = sheet.getRange("A39");
    partIVRange.values = [["四、汇率变动对现金的影响"]];

    ////第五部分
    var partVIRange = sheet.getRange("A40:A41");
    partVIRange.values = [["五、现金及现金等价物净增加额"], ["    加：期初现金及现金等价物余额"]];

    ////第六部分
    var partVIIRange = sheet.getRange("A42");
    partVIIRange.values = [["六、期末现金及现金等价物余额"]];

    sheet.getRange("A43").values = [["逻辑判断"]];

    /////利润
    var profitRange = sheet.getRange("A45:A74");
    var profitValues = [
      [" 1、净利润"],
      ["     加：资产减值损失"],
      ["     固定资产折旧、油气资产折耗、生产性生物资产折旧"],
      ["     无形资产摊销"],
      ["     长期待摊费用的摊销"],
      ["     待摊费用减少（2006版已取消）"],
      ["     预提费用增加（2006版已取消）"],
      ["     处置固定资产、无形资产和其他长期资产的损失"],
      ["     固定资产报废损失"],
      ["     公允价值变动损失"],
      ["     财务费用"],
      ["     投资损失"],
      ["     递延所得税资产的减少"],
      ["     递延所得税负债的增加"],
      ["     存货的减少"],
      ["     经营性应收项目的减少"],
      ["     经营性应付项目的增加"],
      ["     其他"],
      [" 经营活动产生的现金流量净额"],
      ["逻辑判断"],
      [" 2、不涉及现金收支的投资和筹资活动"],
      ["    债务转为资本"],
      ["    一年内到期的可转换公司债券"],
      ["    融资租入固定资产"],
      [" 3、现金及现金等价物净增加情况："],
      ["     现金的期末余额"],
      ["     减：现金的期初余额"],
      ["     加：现金等价物的期末余额"],
      ["     减：现金等价物的期初余额"],
      [" 现金及现金等价物净增加额"]
    ];
    profitRange.values = profitValues;

    sheet.getRange("B9").formulasR1C1 = [["=SUM(R[-3]C:R[-1]C)"]];
    sheet.getRange("B14").formulasR1C1 = [["=SUM(R[-4]C:R[-1]C)"]];
    sheet.getRange("B15").formulasR1C1 = [["=(R[-6]C) - (R[-1]C)"]];
    sheet.getRange("B22").formulasR1C1 = [["=SUM(R[-5]C:R[-1]C)"]];
    sheet.getRange("B27").formulasR1C1 = [["=SUM(R[-4]C:R[-1]C)"]];
    sheet.getRange("B28").formulasR1C1 = [["=(R[-6]C) - (R[-1]C)"]];
    sheet.getRange("B33").formulasR1C1 = [["=SUM(R[-3]C:R[-1]C)"]];
    sheet.getRange("B37").formulasR1C1 = [["=SUM(R[-3]C:R[-1]C)"]];
    sheet.getRange("B38").formulasR1C1 = [["=(R[-5]C) - (R[-1]C)"]];
    sheet.getRange("B42").formulasR1C1 = [["=SUM(R[-2]C:R[-1]C)"]];
    sheet.getRange("B44").values = [["=B3"]];
    sheet.getRange("B45").formulas = [["= 利润表!B26"]];
    sheet.getRange("B63").formulas = [["=B45+SUM(B46:B62)"]];
    sheet.getRange("B64").formulas = [['=IF(B63=B15,"√","×")']];
    sheet.getRange("B74").formulas = [["=B70-B71+B72-B73"]];

    // 设置数字格式
    // sheet.getRange("C6:C28").numberFormat = [["0.00%"]];

    //获取使用范围的行列数
    var useNumRange = sheet.getUsedRange();
    useNumRange.load("columnCount");
    useNumRange.load("rowCount");

    await context.sync().then(function() {
      var useColNum = useNumRange.columnCount; /// 3
      var useRowNum = useNumRange.rowCount; /// 45

      ////横加边框  （固定）
      var rangeBorder = sheet.getRange("A2:B2");
      for (var r = 0; r < useRowNum - 1; r++) {
        rangeBorder.getOffsetRange(r, 0).format.borders.getItem("EdgeBottom").style = "Continuous";
      }

      ////列加边框  （固定）
      var rangeBorder = sheet.getRange("A3:A74");
      for (var c = 0; c <= useColNum; c++) {
        rangeBorder.getOffsetRange(0, c).format.borders.getItem("EdgeLeft").style = "Continuous";
      }
    });

    ///加粗显示
    var firstTitleRanges = sheet.getRanges(
      "A5:B5, A16:B16, A29:B29, A39:B39, A40:A40, A42:B42, A45:B45, A63:B63, A65:B65, A69:B69, A74:B74"
    );
    firstTitleRanges.format.font.bold = true;

    ////居中显示
    var centerShowRanges = sheet.getRanges("A1:B1, A3:B3, A9, A14, A15, A22, A27, A28, A33, A37, A38, A43, A64, B44");
    centerShowRanges.format.horizontalAlignment = "Center";
    centerShowRanges.format.verticalAlignment = "Center";
    centerShowRanges.format.font.bold = true;

    ////逻辑判断
    var logicRange = sheet.getRanges("A43:B43, A64:B64");
    logicRange.format.fill.color = "black";
    logicRange.format.font.color = "white";

    //净额
    var netAmountRange = sheet.getRanges("A15:B15, A28:B28, A38:B38");
    netAmountRange.format.fill.color = "#00b0f0";

    //// 取消行
    sheet.getRange("A50:B51").format.fill.color = "#ffff00";

    // 合并单元格
    sheet.getRange("A3:A4").merge();
    sheet.getRange("B3:B4").merge();

    // 单元格数据格式
    sheet.getRange("B:B").numberFormat = [["#,##0.00"]];
    /////设置行高
    sheet.getUsedRange().format.autofitRows();
    sheet.getUsedRange().format.font.size = 12;
    sheet.getUsedRange().format.font.name = "宋体";
    sheet.getRanges("A1:B1").format.rowHeight = 35;
    sheet.getRange("A1").format.font.size = 24;
    sheet.getUsedRange().format.autofitColumns();

    sheet.activate();
  });
}

//////-------新建利润表------//////
async function setProfitSheet() {
  await Excel.run(async (context) => {
    /////------新建表，若表存在，则删除
    const profitSheetName = "利润表";
    context.workbook.worksheets.getItemOrNullObject(profitSheetName).delete();
    const sheet = context.workbook.worksheets.add(profitSheetName);

    ////// 会计分录和分类账 表
    const mainSheetName = "会计分录和分类账";
    var mainSheet = context.workbook.worksheets.getItem(mainSheetName);
    var expensesTable = mainSheet.tables.getItem("ExpensesTable");
    var headerRange = expensesTable.getHeaderRowRange();

    //第一行
    var firstCell = sheet.getRange("A1");
    firstCell.values = [["利润表"]];
    firstCell.format.rowHeight = 35;
    firstCell.format.font.bold = true;

    // 第二行
    var secondCell = sheet.getRange("A2:B2");
    secondCell.values = [["编制单位", "XX公司"]];

    // 第三、四行   标题行
    var thirdFourRow = sheet.getRange("A3:C4");
    thirdFourRow.values = [["项目", "=资产负债表!B3", "结构分析"], ["", "", "=B3"]];
    thirdFourRow.format.font.bold = true;

    ///// 表侧栏内容
    var profitTableSideRange = sheet.getRange("A5:A45");
    var profitTableSideValues = [
      ["一、营业收入"],
      ["  减：营业成本"],
      ["  营业毛利"],
      ["    营业税金及附加"],
      ["    销售费用"],
      ["    管理费用"],
      ["    财务费用"],
      ["    核心利润"],
      ["    资产减值损失"],
      ["  加：公允价值变动收益（损失以“－”号填列）"],
      ["    投资收益（损失以“－”号填列）"],
      ["    其中：对联营企业和合营企业的投资收益"],
      ["    资产处置收益（损失以“-”号填列）"],
      ["    其他收益"],
      ["二、营业利润（亏损以“－”号填列）"],
      ["  加：营业外收入"],
      ["    其中 ：非流动资产处置利得"],
      ["  减：营业外支出"],
      ["    其中 ：非流动资产处置损失"],
      ["三、利润总额（亏损总额以“－”号填列）"],
      ["  减：所得税费用"],
      ["四、净利润（净亏损以“－”号填列）"],
      ["    （一）持续经营净利润"],
      ["    （二）终止经营净利润"],
      ["五、其他综合收益的税后净额"],
      ["    （一）以后不能重分类进损益的其他综合收益"],
      ["1.重新计量设定受益计划净负债或净资产的变动"],
      ["2.权益法下在被投资单位不能重分类进损益的其他综合收益 中享有的份额"],
      ["……"],
      ["    （二）以后将重分类进损益的其他综合收益"],
      ["1.权益法下在被投资单位以后将重分类进损益的其他综合收益中享有的份额"],
      ["2.可供出售金融资产公允价值变动损益"],
      ["3.持有至到期投资重分类为可供出售金融资产损益"],
      ["4.现金流经套期损益的有效部分"],
      ["5.外币财务报表折算差额"],
      ["6.存货或自用房地产转为以FV计量的投资性房地产"],
      ["……"],
      ["六、综合收益总额"],
      ["七、每股收益"],
      ["    （一）基本每股收益"],
      ["    （二）稀释每股收益"]
    ];
    profitTableSideRange.values = profitTableSideValues;

    //一级标题
    var firstTitleRanges = sheet.getRanges("A7:C7, A12:C12, A24:C24, A26:C26, A29:C29, A42:C42, A43:C43");
    firstTitleRanges.format.font.bold = true;
    //二级标题
    var subTitleRanges = sheet.getRanges("A5:C5, A19:C19, A27:C27, A28:C28, A30:C30, A34:C34, A44:C44, A45:C45");
    subTitleRanges.format.font.bold = true;

    //居中显示行
    var centerRowRange = sheet.getRanges("A1,A3:C4, A7, A12, A33, A41");
    centerRowRange.format.verticalAlignment = "Center";
    centerRowRange.format.horizontalAlignment = "Center";

    //根据单元格名查找发生额合计所在位置
    var startCopyRange = headerRange.findOrNullObject("营业收入", {
      completeMatch: true,
      matchCase: false,
      searchDirection: Excel.SearchDirection.forward
    });
    startCopyRange.load("columnIndex");
    var endCopyRange = headerRange.findOrNullObject("检验", {
      completeMatch: true,
      matchCase: false,
      searchDirection: Excel.SearchDirection.forward
    });
    endCopyRange.load("columnIndex");

    headerRange.load("rowIndex");

    //获取使用范围的行列数
    var useNumRange = sheet.getUsedRange();
    useNumRange.load("columnCount");
    useNumRange.load("rowCount");

    //根据单元格名查找发生额合计所在位置
    var foundRange = mainSheet.getUsedRange().find("发生额合计", {
      completeMatch: true,
      matchCase: false,
      searchDirection: Excel.SearchDirection.forward
    });
    foundRange.load("rowIndex");

    await context.sync().then(function() {
      var sumRowNum = foundRange.rowIndex + 1; ///发生行所在行
      var useColNum = useNumRange.columnCount; /// 3
      var useRowNum = useNumRange.rowCount; /// 45

      /////从 会计分录和分类账表 中调用数据到 利润表格中
      sheet.getRange("B5").formulas = [["='会计分录和分类账'!BH" + sumRowNum]];
      sheet.getRange("B6").formulas = [["='会计分录和分类账'!BI" + sumRowNum]];
      sheet.getRange("B7").formulasR1C1 = [["=R[-2]C-R[-1]C"]];
      sheet.getRange("B8").formulas = [["='会计分录和分类账'!BJ" + sumRowNum]];
      sheet.getRange("B9").formulas = [["='会计分录和分类账'!BL" + sumRowNum]];
      sheet.getRange("B10").formulas = [["='会计分录和分类账'!BK" + sumRowNum]];
      sheet.getRange("B11").formulas = [["='会计分录和分类账'!BM" + sumRowNum]];
      sheet.getRange("B12").formulas = [["=B5-B6-B8-B9-B10-B11"]];
      sheet.getRange("B13").formulas = [["='会计分录和分类账'!BN" + sumRowNum]];
      sheet.getRange("B14").formulas = [["='会计分录和分类账'!BO" + sumRowNum]];
      sheet.getRange("B15").formulas = [["='会计分录和分类账'!BP" + sumRowNum]];

      sheet.getRange("B19").formulas = [["=B5-B6-B8-B9-B10-B11-B13+B14+B15"]];
      sheet.getRange("B20").formulas = [["='会计分录和分类账'!BQ" + sumRowNum]];
      sheet.getRange("B22").formulas = [["='会计分录和分类账'!BR" + sumRowNum]];
      sheet.getRange("B24").formulas = [["=B19+B20-B22"]];
      sheet.getRange("B25").formulas = [["='会计分录和分类账'!BS" + sumRowNum]];
      sheet.getRange("B26").formulasR1C1 = [["=R[-2]C-R[-1]C"]];
      sheet.getRange("B29").formulas = [["=B30+B34"]];
      sheet.getRange("B30").formulasR1C1 = [["=SUM(R[3]C:R[1]C)"]];
      sheet.getRange("B34").formulasR1C1 = [["=SUM(R[7]C:R[1]C)"]];
      sheet.getRange("B42").formulas = [["=B26+B29"]];

      ////// 结构分析 列
      sheet.getRange("C5").values = [["100%"]];

      var startAnalyseCell = sheet.getRange("C6");
      for (var m = 0; m < 23; m++) {
        startAnalyseCell.getOffsetRange(m, 0).formulasR1C1 = [["=RC[-1]/(R[-" + (m + 1) + "]C[-1])"]];
      }

      // 设置数字格式
      sheet.getRange("B:B").numberFormat = [["#,##0.00"]];
      sheet.getRange("C6:C28").numberFormat = [["0.00%"]];

      ////横加边框  （固定）
      var rangeBorder = sheet.getRange("A2:C2");
      for (var r = 0; r < useRowNum - 1; r++) {
        rangeBorder.getOffsetRange(r, 0).format.borders.getItem("EdgeBottom").style = "Continuous";
      }

      ////列加边框  （固定）
      var rangeBorder = sheet.getRange("A3:A45");
      for (var c = 0; c <= useColNum; c++) {
        rangeBorder.getOffsetRange(0, c).format.borders.getItem("EdgeLeft").style = "Continuous";
      }
    });

    /////合并单元格
    sheet.getRange("A1:c1").merge();
    sheet.getRange("B2:c2").merge();
    sheet.getRange("A3:A4").merge();
    sheet.getRange("B3:B4").merge();

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.font.size = 12;
    sheet.getUsedRange().format.font.name = "宋体";
    sheet.getRange("B:B").format.columnWidth = 70;
    sheet.getRange("A1").format.font.size = 24;

    sheet.activate();
  });
}

//////-------新建所有者权益变动表------//////
async function setEquitySheet() {
  await Excel.run(async (context) => {
    /////------新建表，若表存在，则删除
    const equitySheetName = "所有者权益变动表";
    context.workbook.worksheets.getItemOrNullObject(equitySheetName).delete();
    const sheet = context.workbook.worksheets.add(equitySheetName);

    ////// 会计分录和分类账 表
    const mainSheetName = "会计分录和分类账";
    var mainSheet = context.workbook.worksheets.getItem(mainSheetName);

    const balanceSheetName = "资产负债表";
    var balanceSheet = context.workbook.worksheets.getItem(balanceSheetName);

    //第一行
    var firstCell = sheet.getRange("A1");
    firstCell.values = [["所有者权益变动表"]];
    firstCell.format.font.size = 24;
    firstCell.format.font.bold = true;

    // 第二行
    var secondCell = sheet.getRange("A2:B2");
    secondCell.values = [["编制单位", "XX公司"]];

    // 第三、四行   标题行
    var thirdRow = sheet.getRange("A3:B3");
    thirdRow.values = [["项目", '="本年金额-"&资产负债表!B3']];

    var tableTitleRow = sheet.getRange("B4:H4");
    var tableTitleValues = [
      ["实收资本（或股本）", "其他权益工具", "资本公职", "减：库存股", "其他综合收益", "盈余公积", "未分配利润"]
    ];
    tableTitleRow.values = tableTitleValues;

    /////表格侧栏
    var tableSideCol = sheet.getRange("A5:A26");
    var tableSideValues = [
      ["一、上年年末余额"],
      ["加：会计政策变更"],
      ["前期差错更正"],
      ["二、本年年初余额"],
      ["三、本年增减变动金额（减少以“－”号填列）"],
      ["（一）净利润"],
      ["（二）其他综合收益"],
      ["上述（一）和（二）小计"],
      ["（三）所有者投入和减少资本"],
      [" 1.所有者投入资本"],
      ["2.股份支付计入所有者权益的金额"],
      ["3.其他"],
      ["（四）利润分配"],
      ["1.提取盈余公积"],
      ["2.对所有者（或股东）的分配"],
      ["3.其他"],
      ["（五）所有者权益内部结转"],
      ["1.资本公积转增资本（或股本）"],
      ["2.盈余公积转增资本（或股本）"],
      ["3.盈余公积弥补亏损"],
      ["4.其他"],
      [" 四、本年年末余额"]
    ];
    tableSideCol.values = tableSideValues;

    var useNumRange = sheet.getUsedRange();
    useNumRange.load("columnCount");
    useNumRange.load("rowCount");

    useNumRange.format.font.name = "宋体";

    //根据单元格名查找发生额合计所在位置
    var foundRange = mainSheet.getUsedRange().find("发生额合计", {
      completeMatch: true,
      matchCase: false,
      searchDirection: Excel.SearchDirection.forward
    });
    foundRange.load("rowIndex");

    await context.sync().then(function() {
      var sumRowNum = foundRange.rowIndex + 1; ///发生行所在行
      var useColNum = useNumRange.columnCount; /// 8
      var useRowNum = useNumRange.rowCount; /// 26

      // 设置单元格的公式
      //////本年年初余额
      var thisYearBeginStart = sheet.getRange("B8");
      for (var t = 0; t < useColNum - 1; t++) {
        thisYearBeginStart.getOffsetRange(0, t).formulasR1C1 = [["=SUM(R5C:R[-1]C)"]];
      }

      ///净利润
      sheet.getRange("H10").formulas = [["=利润表!B26"]];

      ///其他综合收益
      sheet.getRange("F11").formulas = [["=利润表!B29"]];

      ///上述（一）（二）合计
      var oneAndTwoSumStart = sheet.getRange("B12");
      for (var t = 0; t < useColNum - 1; t++) {
        oneAndTwoSumStart.getOffsetRange(0, t).formulasR1C1 = [["=SUM(R[-2]C:R[-1]C)"]];
      }

      ///（三）所有者投入和减少资本
      var thirdSumStart = sheet.getRange("B13");
      for (var t = 0; t < useColNum - 1; t++) {
        thirdSumStart.getOffsetRange(0, t).formulasR1C1 = [["=SUM(R[1]C:R[3]C)"]];
      }

      sheet.getRange("B14").formulas = [["='会计分录和分类账'!BB" + sumRowNum]];
      sheet.getRange("D14").formulas = [["='会计分录和分类账'!BC" + sumRowNum]];

      ///  （四）利润分配
      var fourSumStart = sheet.getRange("B17");
      for (var t = 0; t < useColNum - 1; t++) {
        fourSumStart.getOffsetRange(0, t).formulasR1C1 = [["=SUM(R[1]C:R[3]C)"]];
      }
      // sheet.getRange("H18").values = [["='会计分录和分类账'!BF82"]];
      // sheet.getRange("H19").values = [["='会计分录和分类账'!BF84"]];
      sheet.getRange("G18").values = [["=-H18"]];
      sheet.getRange("H18:H19").format.fill.color = "red";

      ///    （五）所有者权益内部结转
      var fiveSumStart = sheet.getRange("B21");
      for (var t = 0; t < useColNum - 1; t++) {
        fiveSumStart.getOffsetRange(0, t).formulasR1C1 = [["=SUM(R[1]C:R[4]C)"]];
      }

      ///    四、本年年末余额
      var fiveSumStart = sheet.getRange("B26");
      for (var t = 0; t < useColNum - 3; t++) {
        fiveSumStart.getOffsetRange(0, t).formulasR1C1 = [["=R8C+R12C+R13C-R17C+R21C"]];
      }

      sheet.getRange("G26").formulasR1C1 = [["=R8C+R12C+R13C+R17C+R21C"]];
      sheet.getRange("H26").formulasR1C1 = [["=R8C+R12C+R13C+R17C+R21C"]];

      ////横加边框  （固定）
      var rangeBorder = sheet.getRange("A2:H2");
      for (var r = 0; r < useRowNum - 1; r++) {
        rangeBorder.getOffsetRange(r, 0).format.borders.getItem("EdgeBottom").style = "Continuous";
      }

      ////列加边框  （固定）
      var rangeBorder = sheet.getRange("A3:A26");
      for (var c = 0; c < useColNum; c++) {
        rangeBorder.getOffsetRange(0, c).format.borders.getItem("EdgeRight").style = "Continuous";
      }
    }); ////then() ending

    /////设置样式
    var mainTitleRange = sheet.getRanges("A5, A8, A9, A26");
    mainTitleRange.format.fill.color = "#448844";
    mainTitleRange.format.font.color = "white";
    mainTitleRange.format.font.bold = true;

    var subTitleRange = sheet.getRanges("A12, A13, A17, A21");
    subTitleRange.format.fill.color = "#92d050";
    subTitleRange.format.font.bold = true;

    var sumContentRange = sheet.getRanges("B8:H8, B12:H12, B13:H13, B17:H17, B21:H21,  B26:H26");
    sumContentRange.format.fill.color = "#E1FCCF";
    sumContentRange.format.font.bold = true;

    sheet.getRange("A1").format.horizontalAlignment = "Center";
    sheet.getRange("A1").format.verticalAlignment = "Center";

    /////合并单元格
    sheet.getRange("A1:H1").merge();
    sheet.getRange("B2:H2").merge();
    sheet.getRange("A3:A4").merge();

    ////设置数据格式
    sheet.getRange("B11:H26").numberFormat = [["#,##0.00"]];

    //冻结行
    sheet.freezePanes.freezeRows(1);
    sheet.freezePanes.freezeRows(2);
    sheet.freezePanes.freezeRows(3);
    sheet.freezePanes.freezeRows(4);

    ///设置整体的行高和列宽
    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.rowHeight = 22;
    sheet.getRange("A1").format.rowHeight = 35;

    sheet.activate();
  });
}

//////-------新建资产负债表------//////
async function setBalanceSheet() {
  await Excel.run(async (context) => {
    /////------新建表，若表存在，则删除
    const balanceSheetName = "资产负债表";
    context.workbook.worksheets.getItemOrNullObject(balanceSheetName).delete();
    const sheet = context.workbook.worksheets.add(balanceSheetName);

    //第一行
    const firstCell = sheet.getRange("A1");
    firstCell.values = [["资产负债表     会企01表"]];
    firstCell.format.font.bold = true;
    sheet.getRange("A1:F1").merge();
    // 第二行
    const secondCell = sheet.getRange("A2");
    secondCell.values = [["编制单位：甲公司"]];
    sheet.getRange("A2:F2").merge();
    // 第一与第二行样式设置
    sheet.getRange("A1:F2").format.font.size = 12;
    sheet.getRange("A1:F2").format.rowHeight = 22;

    //////------表格标题栏
    const tableHeaderRowRange = sheet.getRange("A3:F4");
    tableHeaderRowRange.values = [
      ["资产", "2017年", "结构分析", "负债和所有者权益", "2017", "结构分析"],
      ["", "", "2017", "", "", "2017"]
    ];

    ////标题栏样式
    var titleRowFormat = tableHeaderRowRange.format;
    titleRowFormat.fill.color = "black";
    titleRowFormat.font.color = "white";
    titleRowFormat.font.size = 11;
    titleRowFormat.font.bold = true;
    titleRowFormat.font.name = "宋体";
    titleRowFormat.horizontalAlignment = "Center";
    titleRowFormat.verticalAlignment = "Center";
    titleRowFormat.rowHeight = 45;

    ///第五行
    sheet.getRange("A5:F5").values = [["流动资产：", "", "", "流动负债：", "", ""]];

    /////表格主体
    var bodyValues = [
      ["货币资金", "", "", "流动资金", "", ""],
      ["以公允价值计量且其变动计入当期损益的金融资产", "", "", "以公允价值计量且其变动计入当期损益的金融负债", "", ""],
      ["衍生金融资产", "", "", "衍生金融负债", "", ""],
      ["应收票据", "", "", "应付票据", "", ""],
      ["应收账款", "", "", "应付账款", "", ""],
      ["预付款项", "", "", "预收款项", "", ""],
      ["应收利息", "", "", "应付职工薪酬", "", ""],
      ["应收股息", "", "", "应交税费", "", ""],
      ["其他应收款", "", "", "应付利息", "", ""],
      ["存货", "", "", "应付股利", "", ""],
      ["", "", "", "其他应付款", "", ""],
      ["持有待售资产", "", "", "持有待售负债", "", ""],
      ["一年内到期的非流动资产", "", "", "一年内到期的非流动负债", "", ""],
      ["其他流动资产", "", "", "其他流动负债", "", ""],
      ["流动资产合计", "", "", "流动负债合计", "", ""],
      ["非流动资产：", "", "", "非流动负债", "", ""],
      ["可供出售金融资产", "", "", "长期借款", "", ""],
      ["持有至到期投资", "", "", "应付债券", "", ""],
      ["长期应收款", "", "", "其中：优先股", "", ""],
      ["长期股权投资", "", "", "永续债", "", ""],
      ["投资性房地产", "", "", "长期应付款", "", ""],
      ["固定资产", "", "", "专项应付款", "", ""],
      ["在建工程", "", "", "预计负债", "", ""],
      ["工程物质", "", "", "递延收益", "", ""],
      ["固定资产清理", "", "", "递延所得税负债", "", ""],
      ["生产性生物资产", "", "", "其他非流动负债", "", ""],
      ["油气资产", "", "", "非流动负债合计", "", ""],
      ["无形资产", "", "", "负债合计", "", ""],
      ["开发支出", "", "", "所有者权益：", "", ""],
      ["商誉", "", "", "实收资本（或股本）", "", ""],
      ["长期待摊费用", "", "", "其他权益工具", "", ""],
      ["递延所得税资产", "", "", "其中：优先股", "", ""],
      ["其他非流动资产", "", "", "永续股", "", ""],
      ["", "", "", "资本公积", "", ""],
      ["", "", "", "减：库存股", "", ""],
      ["", "", "", "其他综合收益", "", ""],
      ["", "", "", "盈余公积", "", ""],
      ["", "", "", "未分配利润", "", ""],
      ["非流动资产合计", "", "", "所有者权益合计", "", ""],
      ["资产总计", "", "", "负债和所有者权益合计", "", ""]
    ];

    sheet.getRange("A6:F45").values = bodyValues;

    sheet.getRanges("A:A, D:D").format.columnWidth = 150;
    sheet.getRanges("A:A, D:D").format.wrapText = true;

    //小标题
    var subTitleRange = sheet.getRanges("A5, A21, D5, D21, D34");
    subTitleRange.format.fill.color = "#448844";
    subTitleRange.format.font.color = "white";
    subTitleRange.format.font.size = 12;
    subTitleRange.format.font.bold = true;
    subTitleRange.format.font.name = "宋体";
    subTitleRange.format.horizontalAlignment = "Center";
    subTitleRange.format.verticalAlignment = "Center";
    // subTitleRange.format.rowHeight = 45;

    sheet.getRanges("A7, A8, A17, D7, D8, D17, D24, D25, D36, D37, D38").format.fill.color = "#92d050";
    // 文字水平居右显示
    sheet.getRanges("D24:D25, D37:D38").format.horizontalAlignment = "Right";

    // 合计单元格
    var sumRange = sheet.getRanges("A20:F20,D33:F33, A44:F44, A45:F45");
    sumRange.format.font.size = 12;
    sumRange.format.font.bold = true;
    sumRange.format.horizontalAlignment = "Center";
    sumRange.format.font.name = "宋体";

    ////横加边框  （固定）
    var rangeBorder = sheet.getRange("A4:F5");
    for (var r = 0; r < 41; r++) {
      rangeBorder.getOffsetRange(r, 0).format.borders.getItem("EdgeBottom").style = "Continuous";
    }

    ////列加边框  （固定）
    var rangeBorder = sheet.getRange("A5:A45");
    for (var c = 0; c <= 6; c++) {
      rangeBorder.getOffsetRange(0, c).format.borders.getItem("EdgeLeft").style = "Continuous";
    }

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    const sheetName = "会计分录和分类账";
    var mainSheet = context.workbook.worksheets.getItem(sheetName);
    var mainRange = mainSheet.getUsedRange().getLastRow();
    mainRange.load("rowIndex");

    //根据单元格名查找发生额合计所在位置
    var qimoRange = mainSheet.getUsedRange().find("期末余额", {
      completeMatch: true,
      matchCase: false,
      searchDirection: Excel.SearchDirection.forward
    });
    qimoRange.load("rowIndex");

    await context.sync().then(function() {
      // console.log(mainRange.rowIndex);
      var qimoRowNum = qimoRange.rowIndex + 1;
      console.log(qimoRowNum);

      sheet.getRange("B6").formulas = [
        [
          "='会计分录和分类账'!K" +
            qimoRowNum +
            "+'会计分录和分类账'!L" +
            qimoRowNum +
            "+'会计分录和分类账'!M" +
            qimoRowNum
        ]
      ];
      sheet.getRange("B7").formulas = [["='会计分录和分类账'!N" + qimoRowNum]];
      sheet.getRange("B9").formulas = [["='会计分录和分类账'!O" + qimoRowNum]];
      sheet.getRange("B10").formulas = [["='会计分录和分类账'!P" + qimoRowNum]];
      sheet.getRange("B11").formulas = [["='会计分录和分类账'!Q" + qimoRowNum]];
      sheet.getRange("B12").formulas = [["='会计分录和分类账'!R" + qimoRowNum]];
      sheet.getRange("B13").formulas = [["='会计分录和分类账'!S" + qimoRowNum]];
      sheet.getRange("B14").formulas = [["='会计分录和分类账'!T" + qimoRowNum]];
      sheet.getRange("B15").formulas = [["=SUM('会计分录和分类账'!U" + qimoRowNum + ": Y" + qimoRowNum + ")"]];
      sheet.getRange("B20").formulas = [["=SUM(B6:B19)"]];
      sheet.getRange("B22").formulas = [["='会计分录和分类账'!Z" + qimoRowNum]];
      sheet.getRange("B23").formulas = [["='会计分录和分类账'!AA" + qimoRowNum]];
      sheet.getRange("B24").formulas = [["='会计分录和分类账'!AB" + qimoRowNum]];
      sheet.getRange("B25").formulas = [["='会计分录和分类账'!AC" + qimoRowNum]];
      sheet.getRange("B26").formulas = [["='会计分录和分类账'!AD" + qimoRowNum]];
      sheet.getRange("B27").formulas = [["=SUM('会计分录和分类账'!AE" + qimoRowNum + ": AF" + qimoRowNum + ")"]];
      sheet.getRange("B28").formulas = [["='会计分录和分类账'!AG" + qimoRowNum]];
      sheet.getRange("B35").formulas = [["='会计分录和分类账'!AH" + qimoRowNum]];
      sheet.getRange("B36").formulas = [["='会计分录和分类账'!AI" + qimoRowNum]];
      sheet.getRange("B37").formulas = [["='会计分录和分类账'!AJ" + qimoRowNum]];
      sheet.getRange("B44").formulas = [["=SUM(B22:B38)"]];
      sheet.getRange("B45").formulas = [["=B44+B20"]];

      sheet.getRange("E6").formulas = [["='会计分录和分类账'!AL" + qimoRowNum]];
      sheet.getRange("E9").formulas = [["='会计分录和分类账'!AM" + qimoRowNum]];
      sheet.getRange("E10").formulas = [["='会计分录和分类账'!AN" + qimoRowNum]];
      sheet.getRange("E11").formulas = [["='会计分录和分类账'!AO" + qimoRowNum]];
      sheet.getRange("E12").formulas = [["='会计分录和分类账'!AP" + qimoRowNum]];
      sheet.getRange("E13").formulas = [["='会计分录和分类账'!AQ" + qimoRowNum]];
      sheet.getRange("E14").formulas = [["='会计分录和分类账'!AR" + qimoRowNum]];
      sheet.getRange("E15").formulas = [["='会计分录和分类账'!AS" + qimoRowNum]];
      sheet.getRange("E16").formulas = [["='会计分录和分类账'!AT" + qimoRowNum]];
      sheet.getRange("E20").formulas = [["=SUM(E6:E19)"]];
      sheet.getRange("E22").formulas = [["='会计分录和分类账'!AU" + qimoRowNum]];
      sheet.getRange("E23").formulas = [["='会计分录和分类账'!AV" + qimoRowNum]];
      sheet.getRange("E26").formulas = [["='会计分录和分类账'!AW" + qimoRowNum]];
      sheet.getRange("E28").formulas = [["='会计分录和分类账'!AX" + qimoRowNum]];
      sheet.getRange("E29").formulas = [["='会计分录和分类账'!AY" + qimoRowNum]];
      sheet.getRange("E30").formulas = [["='会计分录和分类账'!AZ" + qimoRowNum]];
      sheet.getRange("E32").formulas = [["=SUM(E22:E31)"]];
      sheet.getRange("E33").formulas = [["=E32+E20"]];

      /////TODO 所有者权益
      sheet.getRange("E35").formulas = [["=所有者权益变动表!B26"]];
      sheet.getRange("E39").formulas = [["=所有者权益变动表!D26"]];
      sheet.getRange("E40").formulas = [["=所有者权益变动表!E26"]];
      sheet.getRange("E41").formulas = [["=所有者权益变动表!C26"]];
      sheet.getRange("E42").formulas = [["=所有者权益变动表!G26"]];
      sheet.getRange("E43").formulas = [["=所有者权益变动表!H26"]];
      sheet.getRange("E44").formulas = [["=SUM(E41:E43)+E35+E39-E40"]];
      sheet.getRange("E45").formulas = [["=E44+E33"]];

      sheet.getRange("A47").values = [["报表平衡逻辑判断"]];
      sheet.getRange("B47").formulas = [['=IF(E45=B45,"√","×")']];
      sheet.getRange("A47").format.font.bold = true;
      sheet.getRange("B47").format.font.color = "red";
      sheet.getRange("B47").format.font.bold = true;
    });

    sheet.activate();
    ////转化为表格
    var expensesTable = sheet.tables.add("A3:F45", true);
    expensesTable.name = "BalanceSheetTable";

    ////合并单元格
    sheet.getRange("A3:A4").merge();
    sheet.getRange("B3:B4").merge();
    sheet.getRange("D3:D4").merge();
    sheet.getRange("E3:E4").merge();

    ////设置数据格式
    sheet.getRange("B5:B45").numberFormat = [["#,##0.00"]];
    sheet.getRange("D5:D45").numberFormat = [["#,##0.00"]];
  });
}

//////--------新建会计分录和分账表 ------------------///////
async function setUp() {
  await Excel.run(async (context) => {
    const sheetName = "会计分录和分类账";
    context.workbook.worksheets.getItemOrNullObject(sheetName).delete();
    const sheet = context.workbook.worksheets.add(sheetName);

    /////-------第一栏：表名
    var yearRange = sheet.getRange("A1");
    yearRange.values = [["2017年"]];
    var sheetTitleRange = sheet.getRange("B1");
    sheetTitleRange.values = [["表名"]];
    sheetTitleRange.format.rowHeight = 24;
    sheet.getRange("B1:I1").merge();

    /////-------第二、三栏：平衡验证
    sheet.getRange("D2:D3").values = [["发生额试算平衡结论 "], [" 余额试算平衡结论"]];
    sheet.getRange("G2").formulas = [['=IF(G25=H25,"平衡！","不平衡，错啦！")']];

    // sheet.getRange("G3").formulas = [['=IF(资产负债表!B45=资产负债表!E45,"平衡！","不平衡，错啦！")']];

    ////--------第四栏：建标题栏
    //// 主标题
    const mainTitleRange = sheet.getRange("A4:J4");
    var mainTitleValues = [
      ["序号", "月", "日", "摘要", "一级科目", "次级科目", "借方金额", "贷方金额", "过账", "现金标记"]
    ];
    mainTitleRange.values = mainTitleValues;

    //// 资产余额标题
    const assetTitleRange = sheet.getRange("K4:AJ4");
    var assetTitleValues = [
      [
        "现金",
        "银行存款",
        "其他货币资金",
        "交易性金融资产",
        "应收票据",
        "应收账款",
        "预付账款",
        "应收利息",
        "应收股利",
        "其他应收款",
        "原材料",
        "生产成本",
        "产成品",
        "库存商品",
        "制造费用",
        "可供出售金融资产",
        "持有至到期投资",
        "长期应收款",
        "长期股权投资",
        "投资性房地产",
        "固定资产",
        "累计折旧",
        "在建工程",
        "商誉",
        "长期待摊费用",
        "递延所得税资产"
      ]
    ];
    assetTitleRange.values = assetTitleValues;
    assetTitleRange.getOffsetRange(-1, 0).format.fill.color = "#92d050";
    sheet.getRange("AK4").values = [["+"]];

    //// 负债余额标题
    const liabilityTitleRange = sheet.getRange("AL4:AZ4");
    var liabilityTitleValues = [
      [
        "短期借款",
        "应付票据",
        "应付账款",
        "预收账款",
        "应付职工薪酬",
        "应交税费",
        "应付利息",
        "应付股利",
        "其他应付款",
        "长期借款",
        "应付债券",
        "长期应付款",
        "预计负债",
        "递延收益",
        "递延所得税负债"
      ]
    ];
    liabilityTitleRange.values = liabilityTitleValues;
    liabilityTitleRange.getOffsetRange(-1, 0).format.fill.color = "#92cddc";
    sheet.getRange("BA4").values = [["+"]];

    //// 所有者权益余额
    const equityTitleRange = sheet.getRange("BB4:BF4");
    var equityTitleValues = [["实收资本", "资本公积", "其他综合收益", "盈余公积", "利润分配"]];
    equityTitleRange.values = equityTitleValues;
    equityTitleRange.getOffsetRange(-1, 0).format.fill.color = "#E58CFB";
    sheet.getRange("BG4").values = [["+"]];

    //// 利润
    const profitTitleRange = sheet.getRange("BH4:BT4");
    var profitTitleValues = [
      [
        "营业收入",
        "营业成本",
        "税金及附加",
        "管理费用",
        "营业费用",
        "财务费用",
        "资产减值损失",
        "公允价值变动损益",
        "投资收益",
        "营业外收入",
        "营业外支出",
        "所得税费用",
        "本年利润"
      ]
    ];
    profitTitleRange.values = profitTitleValues;
    profitTitleRange.getOffsetRange(-1, 0).format.fill.color = "#EFDE98";
    sheet.getRange("BU4").values = [["检验"]];

    // 自动换行
    sheet.getRange("J4:BU4").format.columnWidth = 60;
    sheet.getRange("J4:BU4").format.wrapText = true;

    ////--------初期余额
    sheet.getRange("D5").values = [["初期余额"]];
    sheet.getRange("I5").formulas = [['=IF(SUM(K5:AJ5)=(SUM(AL5:AZ5)+SUM(BB5:BF5)),"对！","错！")']];
    sheet.getRange("D5").format.font.bold = true;

    ////--------发生额合计
    var actualAmountRowRange = sheet.getRange("A25:I25");
    actualAmountRowRange.values = [["", "", "", "发生额合计", "", "", "=SUM(R6C:R[-1]C)", "=SUM(R6C:R[-1]C)", ""]];
    actualAmountRowRange.format.font.bold = true;

    /////
    var otherRange = sheet.getRange("D30:D32");
    otherRange.values = [["资产余额"], ["负债余额"], ["所有者权益余额"]];
    otherRange.format.font.bold = true;
    sheet.getRange("D30").format.fill.color = "#92d050";
    sheet.getRange("D31").format.fill.color = "#92cddc";
    sheet.getRange("D32").format.fill.color = "#E58CFB";

    sheet.getRange("E30").formulas = [["=SUM(K28:AE28)+AF28"]];
    sheet.getRange("E31").formulas = [["=SUM(AL28:AZ28)"]];
    sheet.getRange("E32").formulas = [["=SUM(BB28:BF28)"]];

    //冻结表头
    toFreezeColumns();

    /////------自动调整单元格高宽
    sheet.getRange("A:F").format.autofitColumns();
    sheet.getRange("A:F").format.autofitRows();

    //设置样式
    var titleStyleRange = sheet
      .getUsedRange()
      .getLastColumn()
      .load("columnIndex");

    var kRowNumRange = sheet.getRange("k:k").load("columnIndex"); //10

    /////--------数据验证 加载有几列
    // 资产余额
    assetTitleRange.load("columnCount");
    // 负债余额
    liabilityTitleRange.load("columnCount");
    // 所有者权益余额
    equityTitleRange.load("columnCount");
    // 利润
    profitTitleRange.load("columnCount");

    //根据单元格名查找发生额合计所在位置
    var foundRange = sheet.getUsedRange().find("发生额合计", {
      completeMatch: true,
      matchCase: false,
      searchDirection: Excel.SearchDirection.forward
    });
    foundRange.load("rowIndex");

    ////----同步----------------------------------------------/////
    await context.sync().then(function() {
      ////k的序号
      var kRowNum = kRowNumRange.columnIndex; /// 10
      ////BU的序号
      var buRowNum = titleStyleRange.columnIndex; ///
      /// k --BU列数
      var totalRow = buRowNum - kRowNum;
      ///发生额所在行
      var occurSumRow = foundRange.rowIndex; /// 24

      /////-------检测列 =----------------/////
      const checkCell = sheet.getCell(4, buRowNum);
      for (var i = 0; i <= occurSumRow - 5; i++) {
        checkCell.getOffsetRange(i, 0).formulasR1C1 = [["=SUM(RC10:RC[-1])"]];
      }

      ////--------发生额合计--------------/////
      ///表示该列的第5个单元格至当前单元格上一行同列求和
      const actualAmountCell = sheet.getRange("k25");
      for (var i = 0; i <= totalRow; i++) {
        actualAmountCell.getOffsetRange(0, i).formulasR1C1 = [["=SUM(R6C:R[-1]C)"]];
      }

      /////--------------期末余额-----------//////
      sheet.getRange("D28").values = [["期末余额"]];
      ///表示该列的第5个单元格 + 该列发生额合计
      const finalBalanceCell = sheet.getRange("k28");
      var totalRow = titleStyleRange.columnIndex - kRowNum;
      for (var i = 0; i <= totalRow; i++) {
        finalBalanceCell.getOffsetRange(0, i).formulasR1C1 = [["=SUM(R5C, R25C)"]];
      }

      //////---------期末验证----------/////////
      // 资产余额
      var assetRowNum = assetTitleRange.columnCount;
      // 负债余额
      var liabilityRowNum = liabilityTitleRange.columnCount;
      // 所有者权益余额
      var equityRowNum = equityTitleRange.columnCount;
      // 利润
      var profitRowNum = profitTitleRange.columnCount;

      dataValidation(sheet.getRange("K3"), assetRowNum);
      dataValidation(sheet.getRange("AL3"), liabilityRowNum);
      dataValidation(sheet.getRange("BB3"), equityRowNum);
      dataValidation(sheet.getRange("BH3"), profitRowNum);

      // 设置数据格式
      sheet.getRange("G:H").numberFormat = [["#,##0.00"]];
      sheet.getRange("K:BU").numberFormat = [["#,##,##,##0.00"]];

      // 设置第一行样式
      var titleRowFormat = sheet.getRangeByIndexes(3, 0, 1, titleStyleRange.columnIndex + 1).format;
      titleRowFormat.fill.color = "black";
      titleRowFormat.font.color = "white";
      titleRowFormat.font.size = 11;
      titleRowFormat.font.bold = true;
      titleRowFormat.font.name = "宋体";
      titleRowFormat.horizontalAlignment = "Center";
      titleRowFormat.verticalAlignment = "Center";
      titleRowFormat.rowHeight = 35;
      // 设置表格标题栏的边框样式
      for (var i = 0; i < titleStyleRange.columnIndex; i++) {
        var borderFormat = sheet.getCell(3, i).format.borders;
        borderFormat.getItem("EdgeRight").color = "white";
        borderFormat.getItem("EdgeRight").weight = "Medium";
      }

      sheet.getRange("A22:BU24").format.fill.color = "#ccc0da";
    }); //// ---同步 ending

    ////----------------转化为表格-------------/////
    var expensesTable = sheet.tables.add("A4:BU25", true);
    expensesTable.name = "ExpensesTable";
    sheet.activate();

    //////-----过账
    var startPostCell = sheet.getRange("I6");
    var postcolNum = 19;
    postValidation(startPostCell, postcolNum);

    // 给单元格加下拉框
    await context.sync().then(function() {
      addCellDropDownRange();
    });

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();
  });
} /////新建表 ending

//////---------期末余额验证
async function dataValidation(startCell, rowNum) {
  for (var k = 0; k < rowNum; k++) {
    startCell.getOffsetRange(0, k).formulasR1C1 = [['=IF(R28C < 0, "余额为负啦！", "正常")']];
  }
}

//////---------过账
async function postValidation(startPostCell, colNum) {
  for (var m = 0; m < colNum; m++) {
    startPostCell.getOffsetRange(m, 0).formulasR1C1 = [
      ['=IF( OR( (SUM(RC[-2]:RC[-1]) = SUM(RC11:RC72)), (SUM(RC[-2]:RC[-1]) = -SUM(RC11:RC72)) ),"√", "过账错误！")']
    ];
  }
}

//////------------------------借方减贷方----------------------------------------////
// 默认借方金额和贷方金额为 G:7和H:8
async function borrowersToLenders() {
  await Excel.run(async (context) => {
    const sheetName = "会计分录和分类账";
    var sheet = context.workbook.worksheets.getItem(sheetName);
    //点击获取单元格的位置
    const activeRange = context.workbook.getActiveCell();
    activeRange.formulasR1C1 = [["=RC7-RC8"]];
    await context.sync();
  });
}

//////------------------------贷方减借方----------------------------------------////
// 默认借方金额和贷方金额为 G:7和H:8
async function lendersToBorrowers() {
  await Excel.run(async (context) => {
    const sheetName = "会计分录和分类账";
    var sheet = context.workbook.worksheets.getItem(sheetName);
    //点击获取单元格的位置
    const activeRange = context.workbook.getActiveCell();
    activeRange.formulasR1C1 = [["=RC8-RC7"]];
    await context.sync();
  });
}

//////---------------给单元格加下拉框--------------////
async function addCellDropDownRange() {
  await Excel.run(async (context) => {
    const mainSheetName = "会计分录和分类账";
    var mainSheet = context.workbook.worksheets.getItem(mainSheetName);
    var expensesTable = mainSheet.tables.getItem("ExpensesTable");
    var headerRange = expensesTable.getHeaderRowRange();

    var cellDropDownRange = mainSheet.getRange("E6:E24");

    ////根据单元格名查找 “现金标记” 所在位置
    var xianjinCell = headerRange.find("现金标记", {
      completeMatch: true,
      matchCase: false,
      searchDirection: Excel.SearchDirection.forward
    });
    //开始单元格
    var startCell = xianjinCell.getOffsetRange(0, 1);
    startCell.load("address");

    ////根据单元格名查找 “检验” 所在位置
    var jianyanCell = headerRange.find("检验", {
      completeMatch: true,
      matchCase: false,
      searchDirection: Excel.SearchDirection.forward
    });
    // 结束单元格
    var endCell = jianyanCell.getOffsetRange(0, -1);
    endCell.load("address");

    await context.sync().then(function() {
      ///获取开始单元格所在列的符号
      var start = startCell.address.replace(mainSheetName + "!", "").match("[a-zA-Z]")[0];
      ///获取结束单元格所在列的符号
      var end = endCell.address.replace(mainSheetName + "!", "").match("[a-zA-Z]+")[0];

      //// 加下拉框
      cellDropDownRange.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: "=" + mainSheetName + "!$" + start + "$4:$" + end + "$4"
        }
      };
    });
  });
}

//////------------------------冻结第一行和特定列------------------------------////
async function toFreezeColumns() {
  await Excel.run(async (context) => {
    const sheetName = "会计分录和分类账";
    const sheet = context.workbook.worksheets.getItem(sheetName);
    //冻结第一行
    sheet.freezePanes.freezeRows(1);
    sheet.freezePanes.freezeRows(2);
    sheet.freezePanes.freezeRows(3);
    sheet.freezePanes.freezeRows(4);
    //冻结第1~8列
    sheet.freezePanes.freezeColumns(1);
    sheet.freezePanes.freezeColumns(2);
    sheet.freezePanes.freezeColumns(3);
    sheet.freezePanes.freezeColumns(4);
    sheet.freezePanes.freezeColumns(5);
    sheet.freezePanes.freezeColumns(6);
    sheet.freezePanes.freezeColumns(7);
    sheet.freezePanes.freezeColumns(8);
    sheet.freezePanes.freezeColumns(9);

    await context.sync();
  });
}

//////--------------------示例--------------------///////
// async function run() {
//   await Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getActiveWorksheet();

//     console.log("Your code goes here");

//     await context.sync();
//   });
// }

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

(function(){
  'use strict';

  const Core=window.ReconcileCore;
  const state={order:null,payment:null,result:null,filter:'全部',search:''};
  const $=selector=>document.querySelector(selector);
  const elements={
    orderFile:$('#order-file'),paymentFile:$('#payment-file'),orderDrop:$('#order-drop'),paymentDrop:$('#payment-drop'),
    orderFileName:$('#order-file-name'),paymentFileName:$('#payment-file-name'),orderFileMeta:$('#order-file-meta'),paymentFileMeta:$('#payment-file-meta'),
    orderSheet:$('#order-sheet'),paymentSheet:$('#payment-sheet'),orderHeader:$('#order-header-row'),paymentHeader:$('#payment-header-row'),
    orderKey:$('#order-key'),paymentKey:$('#payment-key'),orderAmount:$('#order-amount'),paymentAmount:$('#payment-amount'),
    mappingBlock:$('#mapping-block'),runButton:$('#run-button'),exportButton:$('#export-button'),empty:$('#empty-state'),results:$('#results'),
    error:$('#error-box'),errorMessage:$('#error-message'),body:$('#result-body'),search:$('#result-search'),toast:$('#toast')
  };

  function showToast(message){elements.toast.textContent=message;elements.toast.classList.add('show');clearTimeout(showToast.timer);showToast.timer=setTimeout(()=>elements.toast.classList.remove('show'),2800)}
  function showError(message){elements.error.hidden=false;elements.errorMessage.textContent=message;elements.empty.hidden=true;elements.results.hidden=true;showToast(message)}
  function clearError(){elements.error.hidden=true;elements.errorMessage.textContent=''}
  function formatFileSize(bytes){return bytes<1024*1024?`${Math.max(1,Math.round(bytes/1024))} KB`:`${(bytes/1024/1024).toFixed(1)} MB`}
  function formatMoney(value){return Number(value||0).toLocaleString('zh-CN',{minimumFractionDigits:2,maximumFractionDigits:2})}
  function escapeHTML(value){return String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]))}

  async function parseFile(file,kind){
    if(!window.XLSX) throw new Error('Excel 处理组件未加载，请重新打开页面');
    if(file.size>35*1024*1024) throw new Error('单个文件不能超过 35 MB');
    const buffer=await file.arrayBuffer();
    let workbook;
    try{workbook=XLSX.read(buffer,{type:'array',cellDates:true,dense:true})}catch(error){throw new Error(`无法读取 ${file.name}，请确认文件未损坏且格式正确`)}
    if(!workbook.SheetNames.length) throw new Error(`${file.name} 中没有可读取的工作表`);
    const item={file,workbook,sheetName:workbook.SheetNames[0],headers:[],rows:[],matrix:[],truncated:false};
    state[kind]=item;
    fillSheetSelect(kind);
    refreshDataset(kind);
    updateUploadCard(kind);
    updateReadyState();
  }

  function sheetMatrix(item){
    const worksheet=item.workbook.Sheets[item.sheetName];
    return XLSX.utils.sheet_to_json(worksheet,{header:1,defval:'',raw:false,blankrows:true});
  }

  function refreshDataset(kind){
    const item=state[kind];if(!item)return;
    const headerInput=kind==='order'?elements.orderHeader:elements.paymentHeader;
    item.matrix=sheetMatrix(item);
    const parsed=Core.rowsFromMatrix(item.matrix,headerInput.value);
    item.headers=parsed.headers;item.rows=parsed.rows;item.truncated=parsed.truncated;
    fillColumnSelects(kind);
    updateUploadCard(kind);
  }

  function fillSheetSelect(kind){
    const item=state[kind];const select=kind==='order'?elements.orderSheet:elements.paymentSheet;
    select.innerHTML=item.workbook.SheetNames.map(name=>`<option value="${escapeHTML(name)}">${escapeHTML(name)}</option>`).join('');
    select.value=item.sheetName;
  }

  function guessHeader(headers,words,fallback=0){
    for(const word of words){const found=headers.find(header=>header.toLowerCase().includes(word.toLowerCase()));if(found)return found}
    return headers[Math.min(fallback,Math.max(0,headers.length-1))]||'';
  }

  function fillColumnSelects(kind){
    const item=state[kind];const keySelect=kind==='order'?elements.orderKey:elements.paymentKey;const amountSelect=kind==='order'?elements.orderAmount:elements.paymentAmount;
    const html=item.headers.map(header=>`<option value="${escapeHTML(header)}">${escapeHTML(header)}</option>`).join('');keySelect.innerHTML=html;amountSelect.innerHTML=html;
    keySelect.value=guessHeader(item.headers,['订单号','订单编号','流水号','单号','编号','id']);
    amountSelect.value=guessHeader(item.headers,kind==='order'?['应收金额','订单金额','应收','金额','合计']:['实收金额','收款金额','到账金额','实收','金额','合计'],item.headers.length-1);
  }

  function updateUploadCard(kind){
    const item=state[kind];if(!item)return;
    const card=kind==='order'?elements.orderDrop:elements.paymentDrop;const name=kind==='order'?elements.orderFileName:elements.paymentFileName;const meta=kind==='order'?elements.orderFileMeta:elements.paymentFileMeta;
    card.classList.add('loaded');name.textContent=item.file.name;meta.textContent=`${item.sheetName} · ${item.rows.length.toLocaleString('zh-CN')} 行 · ${formatFileSize(item.file.size)}${item.truncated?' · 已截取前10万行':''}`;
  }

  function updateReadyState(){
    const ready=Boolean(state.order&&state.payment);elements.mappingBlock.classList.toggle('disabled',!ready);elements.runButton.disabled=!ready;
  }

  function runReconciliation(){
    clearError();
    try{
      refreshDataset('order');refreshDataset('payment');
      const config={orderKey:elements.orderKey.value,orderAmount:elements.orderAmount.value,paymentKey:elements.paymentKey.value,paymentAmount:elements.paymentAmount.value,tolerance:$('#tolerance').value};
      if(!config.orderKey||!config.orderAmount||!config.paymentKey||!config.paymentAmount) throw new Error('请完整选择匹配编号和金额字段');
      state.result=Core.reconcile(state.order.rows,state.payment.rows,config);state.result.config=config;state.filter='全部';state.search='';elements.search.value='';
      renderMetrics();renderCounts();renderTable();elements.empty.hidden=true;elements.results.hidden=false;elements.exportButton.disabled=false;
      document.querySelectorAll('#filter-tabs button').forEach(button=>button.classList.toggle('active',button.dataset.filter==='全部'));
      showToast(`对账完成：发现 ${state.result.issues} 个需要处理的编号`);
    }catch(error){showError(error.message);elements.exportButton.disabled=true}
  }

  function renderMetrics(){
    $('#metric-total').textContent=state.result.total.toLocaleString('zh-CN');
    $('#metric-matched').textContent=state.result.matched.toLocaleString('zh-CN');
    $('#metric-issues').textContent=state.result.issues.toLocaleString('zh-CN');
    const sign=state.result.totalDifference>0?'+':'';$('#metric-difference').textContent=`${sign}¥${formatMoney(state.result.totalDifference)}`;
  }

  function renderCounts(){
    const counts=state.result.counts;$('#count-all').textContent=state.result.total;$('#count-matched').textContent=counts[Core.STATUS.MATCHED]||0;$('#count-missing').textContent=counts[Core.STATUS.MISSING]||0;$('#count-diff').textContent=counts[Core.STATUS.DIFFERENCE]||0;$('#count-extra').textContent=counts[Core.STATUS.EXTRA]||0;$('#count-duplicate').textContent=counts[Core.STATUS.DUPLICATE]||0;
  }

  function filteredRows(){
    const search=Core.normalizeKey(state.search);return state.result.results.filter(row=>(state.filter==='全部'||row.status===state.filter)&&(!search||Core.normalizeKey(row.key).includes(search)));
  }

  function statusClass(status){if(status===Core.STATUS.MATCHED)return'good';if(status===Core.STATUS.DIFFERENCE||status===Core.STATUS.MISSING)return'bad';return'issue'}
  function renderTable(){
    const rows=filteredRows();const preview=rows.slice(0,500);
    elements.body.innerHTML=preview.map(row=>{const diffClass=row.difference>0?'difference-positive':row.difference<0?'difference-negative':'';const sign=row.difference>0?'+':'';return `<tr><td><span class="status-pill ${statusClass(row.status)}">${escapeHTML(row.status)}</span></td><td>${escapeHTML(row.key)}</td><td class="money">¥${formatMoney(row.orderAmount)}</td><td class="money">¥${formatMoney(row.paymentAmount)}</td><td class="money ${diffClass}">${sign}¥${formatMoney(row.difference)}</td><td>${row.orderCount}${row.orderCount>1?' · 重复':''}</td><td>${row.paymentCount}${row.paymentCount>1?' · 重复':''}</td></tr>`}).join('');
    $('#table-summary').textContent=`当前筛选 ${rows.length.toLocaleString('zh-CN')} 条`;
  }

  function safeRows(rows){return rows.map(row=>Object.fromEntries(Object.entries(row).filter(([key])=>key!=='_sourceRow').map(([key,value])=>[key,Core.safeSpreadsheetValue(value)])))}
  function addJsonSheet(workbook,name,rows,headers,widths){
    const safe=safeRows(rows);const worksheet=XLSX.utils.json_to_sheet(safe,{header:headers});worksheet['!autofilter']={ref:XLSX.utils.encode_range({s:{r:0,c:0},e:{r:Math.max(0,safe.length),c:headers.length-1}})};worksheet['!cols']=widths.map(width=>({wch:width}));XLSX.utils.book_append_sheet(workbook,worksheet,name);
  }

  async function exportWorkbook(){
    if(!state.result)return;
    try{
      const workbook=XLSX.utils.book_new();const generated=new Date();
      const overview=[['订单与收款对账报告'],[],['生成时间',generated.toLocaleString('zh-CN')],['订单文件',state.order.file.name],['收款文件',state.payment.file.name],['订单匹配字段',state.result.config.orderKey],['收款匹配字段',state.result.config.paymentKey],['允许误差',state.result.tolerance],[],['指标','数量或金额'],['唯一编号',state.result.total],['核对一致',state.result.matched],['需要处理',state.result.issues],['未收到款',state.result.counts[Core.STATUS.MISSING]||0],['额外收款',state.result.counts[Core.STATUS.EXTRA]||0],['金额不一致',state.result.counts[Core.STATUS.DIFFERENCE]||0],['重复编号',state.result.counts[Core.STATUS.DUPLICATE]||0],['订单总额',state.result.totalOrder],['收款总额',state.result.totalPayment],['总金额差（实收-应收）',state.result.totalDifference],[],['数据检查','结果'],['订单表空白编号行',state.result.blankOrderKeys.join('、')||'无'],['收款表空白编号行',state.result.blankPaymentKeys.join('、')||'无']];
      const overviewSheet=XLSX.utils.aoa_to_sheet(overview);overviewSheet['!cols']=[{wch:24},{wch:42}];XLSX.utils.book_append_sheet(workbook,overviewSheet,'对账概览');
      const detailHeaders=['对账状态','匹配编号','订单笔数','订单金额','收款笔数','收款金额','差额（实收-应收）','订单表行号','收款表行号','检查说明'];
      const detailRows=state.result.results.map(row=>({'对账状态':row.status,'匹配编号':Core.safeSpreadsheetValue(row.key),'订单笔数':row.orderCount,'订单金额':row.orderAmount,'收款笔数':row.paymentCount,'收款金额':row.paymentAmount,'差额（实收-应收）':row.difference,'订单表行号':row.orderRows,'收款表行号':row.paymentRows,'检查说明':row.flags}));
      addJsonSheet(workbook,'对账明细',detailRows,detailHeaders,[14,20,10,14,10,14,18,16,16,24]);
      addJsonSheet(workbook,'异常明细',detailRows.filter(row=>row['对账状态']!==Core.STATUS.MATCHED),detailHeaders,[14,20,10,14,10,14,18,16,16,24]);
      const orderSourceRows=state.order.rows.map(({_sourceRow,...row})=>({'原文件行号':_sourceRow,...row}));
      const paymentSourceRows=state.payment.rows.map(({_sourceRow,...row})=>({'原文件行号':_sourceRow,...row}));
      addJsonSheet(workbook,'订单原表',orderSourceRows,['原文件行号',...state.order.headers],[12,...state.order.headers.map(()=>14)]);
      addJsonSheet(workbook,'收款原表',paymentSourceRows,['原文件行号',...state.payment.headers],[12,...state.payment.headers.map(()=>14)]);
      const guide=[['使用说明'],['1. 对账概览：查看数量、金额合计及空白编号。'],['2. 对账明细：查看全部唯一编号的核对结果。'],['3. 异常明细：只保留需要人工处理的记录。'],['4. 原表：保留本次读取的数据，“原文件行号”方便回到源文件复核。'],[],['状态说明'],['核对一致：两张表都有该编号，金额差在允许误差内，且没有重复。'],['未收到款：订单表有编号，收款表没有。'],['仅收款表有：收款表有编号，订单表没有。'],['金额不一致：两张表都有编号，但金额差超出允许误差。'],['重复编号：同一编号在任意一张表中出现多行，程序已按编号合计。']];const guideSheet=XLSX.utils.aoa_to_sheet(guide);guideSheet['!cols']=[{wch:88}];XLSX.utils.book_append_sheet(workbook,guideSheet,'使用说明');
      const date=`${generated.getFullYear()}${String(generated.getMonth()+1).padStart(2,'0')}${String(generated.getDate()).padStart(2,'0')}_${String(generated.getHours()).padStart(2,'0')}${String(generated.getMinutes()).padStart(2,'0')}`;
      const fileName=`订单收款对账报告_${date}.xlsx`;
      if(window.desktopBridge?.saveWorkbook){
        const bytes=XLSX.write(workbook,{bookType:'xlsx',type:'array',compression:true});
        const result=await window.desktopBridge.saveWorkbook(fileName,bytes);
        showToast(result.saved?'Excel 报告已保存':'已取消保存');
      }else{
        XLSX.writeFile(workbook,fileName,{compression:true});showToast('Excel 报告已生成');
      }
    }catch(error){showError(`导出失败：${error.message}`)}
  }

  function workbookFile(name,sheetName,rows){const workbook=XLSX.utils.book_new();const sheet=XLSX.utils.json_to_sheet(rows);sheet['!cols']=Object.keys(rows[0]).map(()=>({wch:16}));XLSX.utils.book_append_sheet(workbook,sheet,sheetName);const bytes=XLSX.write(workbook,{bookType:'xlsx',type:'array',compression:true});return new File([bytes],name,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})}
  async function loadDemo(){
    const orders=[{'订单号':'DD-1001','订单日期':'2026-11-01','客户':'远山商贸','应收金额':1280},{'订单号':'DD-1002','订单日期':'2026-11-01','客户':'星海工作室','应收金额':860},{'订单号':'DD-1003','订单日期':'2026-11-02','客户':'青禾门店','应收金额':500},{'订单号':'DD-1004','订单日期':'2026-11-02','客户':'晨光商店','应收金额':1200},{'订单号':'DD-1005','订单日期':'2026-11-03','客户':'拾光摄影','应收金额':680},{'订单号':'DD-1005','订单日期':'2026-11-03','客户':'拾光摄影','应收金额':320},{'订单号':'DD-1006','订单日期':'2026-11-04','客户':'云杉科技','应收金额':1500}];
    const payments=[{'订单编号':'DD-1001','收款日期':'2026-11-01','支付渠道':'银行转账','实收金额':1280},{'订单编号':'DD-1002','收款日期':'2026-11-01','支付渠道':'微信','实收金额':800},{'订单编号':'DD-1003','收款日期':'2026-11-02','支付渠道':'支付宝','实收金额':500},{'订单编号':'DD-1005','收款日期':'2026-11-03','支付渠道':'银行转账','实收金额':1000},{'订单编号':'DD-1006','收款日期':'2026-11-04','支付渠道':'银行转账','实收金额':1000},{'订单编号':'DD-1006','收款日期':'2026-11-05','支付渠道':'微信','实收金额':500},{'订单编号':'DD-1008','收款日期':'2026-11-05','支付渠道':'微信','实收金额':360}];
    const orderFile=workbookFile('演示订单表.xlsx','订单明细',orders);const paymentFile=workbookFile('演示收款表.xlsx','收款明细',payments);
    await parseFile(orderFile,'order');await parseFile(paymentFile,'payment');showToast('演示数据已加载，可以开始对账');
  }

  function reset(){state.order=null;state.payment=null;state.result=null;state.filter='全部';state.search='';elements.orderFile.value='';elements.paymentFile.value='';elements.orderDrop.classList.remove('loaded');elements.paymentDrop.classList.remove('loaded');elements.orderFileName.textContent='选择 XLSX / CSV 文件';elements.paymentFileName.textContent='选择 XLSX / CSV 文件';elements.orderFileMeta.textContent='包含订单编号和应收金额';elements.paymentFileMeta.textContent='包含订单编号和实收金额';elements.mappingBlock.classList.add('disabled');elements.runButton.disabled=true;elements.exportButton.disabled=true;elements.empty.hidden=false;elements.results.hidden=true;clearError();['#metric-total','#metric-matched','#metric-issues','#metric-difference'].forEach(selector=>$(selector).textContent='—');showToast('已清空，可以重新选择文件')}

  function copyRequirements(){const text=`对账工具定制需求清单\n1. 订单表样例（请脱敏）：\n2. 收款表样例（请脱敏）：\n3. 两张表用于匹配的字段：\n4. 两张表的金额字段：\n5. 是否存在拆单、合并付款、退款或手续费：\n6. 允许的金额误差：\n7. 需要导出的字段和公司模板：\n8. 每次大约处理多少行：`;
    navigator.clipboard?.writeText(text).then(()=>showToast('客户需求清单已复制')).catch(()=>showToast('浏览器未允许复制，请从使用说明中手动记录'))}

  elements.orderFile.addEventListener('change',event=>event.target.files[0]&&parseFile(event.target.files[0],'order').catch(error=>showError(error.message)));
  elements.paymentFile.addEventListener('change',event=>event.target.files[0]&&parseFile(event.target.files[0],'payment').catch(error=>showError(error.message)));
  [['order',elements.orderDrop],['payment',elements.paymentDrop]].forEach(([kind,drop])=>{['dragenter','dragover'].forEach(type=>drop.addEventListener(type,event=>{event.preventDefault();drop.classList.add('drag')}));['dragleave','drop'].forEach(type=>drop.addEventListener(type,event=>{event.preventDefault();drop.classList.remove('drag')}));drop.addEventListener('drop',event=>{const file=event.dataTransfer.files[0];if(file)parseFile(file,kind).catch(error=>showError(error.message))})});
  elements.orderSheet.addEventListener('change',()=>{state.order.sheetName=elements.orderSheet.value;refreshDataset('order')});elements.paymentSheet.addEventListener('change',()=>{state.payment.sheetName=elements.paymentSheet.value;refreshDataset('payment')});
  elements.orderHeader.addEventListener('change',()=>state.order&&refreshDataset('order'));elements.paymentHeader.addEventListener('change',()=>state.payment&&refreshDataset('payment'));
  elements.runButton.addEventListener('click',runReconciliation);elements.exportButton.addEventListener('click',exportWorkbook);$('#reset-button').addEventListener('click',reset);$('#load-demo').addEventListener('click',()=>loadDemo().catch(error=>showError(error.message)));$('#empty-demo').addEventListener('click',()=>loadDemo().catch(error=>showError(error.message)));
  $('#filter-tabs').addEventListener('click',event=>{const button=event.target.closest('button[data-filter]');if(!button)return;state.filter=button.dataset.filter;document.querySelectorAll('#filter-tabs button').forEach(item=>item.classList.toggle('active',item===button));renderTable()});
  elements.search.addEventListener('input',()=>{state.search=elements.search.value;renderTable()});
  $('#copy-requirements').addEventListener('click',copyRequirements);const dialog=$('#guide-dialog');$('#open-guide').addEventListener('click',()=>dialog.showModal());$('#close-guide').addEventListener('click',()=>dialog.close());dialog.addEventListener('click',event=>{if(event.target===dialog)dialog.close()});
})();

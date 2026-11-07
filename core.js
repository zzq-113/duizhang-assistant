(function(global){
  'use strict';

  const MAX_ROWS = 100000;
  const STATUS = {
    MATCHED: '核对一致',
    MISSING: '未收到款',
    EXTRA: '仅收款表有',
    DIFFERENCE: '金额不一致',
    DUPLICATE: '重复编号'
  };

  function normalizeKey(value){
    return String(value ?? '').trim().replace(/\s+/g, '').toLocaleUpperCase('zh-CN');
  }

  function parseAmount(value){
    if(typeof value === 'number'){
      if(!Number.isFinite(value)) throw new Error(`金额不是有效数字：${value}`);
      return value;
    }
    const cleaned = String(value ?? '')
      .replace(/[￥¥$,，\s]/g, '')
      .replace(/^\((.*)\)$/, '-$1')
      .trim();
    if(cleaned === '') return 0;
    const number = Number(cleaned);
    if(!Number.isFinite(number)) throw new Error(`金额不是有效数字：${value}`);
    return number;
  }

  function roundMoney(value){
    return Math.round((value + Number.EPSILON) * 100) / 100;
  }

  function makeHeaders(values){
    const counts = new Map();
    return values.map((value,index)=>{
      const base = String(value ?? '').trim() || `未命名列${index+1}`;
      const count = (counts.get(base) || 0) + 1;
      counts.set(base,count);
      return count === 1 ? base : `${base} (${count})`;
    });
  }

  function rowsFromMatrix(matrix, headerRow){
    const rowIndex = Math.max(0, Number(headerRow || 1) - 1);
    if(!Array.isArray(matrix) || !matrix[rowIndex]) throw new Error(`第 ${rowIndex+1} 行没有找到表头`);
    const headers = makeHeaders(matrix[rowIndex]);
    const rows = matrix.slice(rowIndex+1,rowIndex+1+MAX_ROWS)
      .map((row,offset)=>{
        const record = {_sourceRow: rowIndex + offset + 2};
        headers.forEach((header,index)=>{record[header] = row[index] ?? ''});
        return record;
      })
      .filter(record=>headers.some(header=>String(record[header] ?? '').trim() !== ''));
    return {headers,rows,truncated:matrix.length-rowIndex-1>MAX_ROWS};
  }

  function groupRows(rows,keyColumn,amountColumn,sourceLabel){
    const groups = new Map();
    const blankKeys = [];
    rows.forEach(row=>{
      const rawKey = row[keyColumn];
      const key = normalizeKey(rawKey);
      if(!key){blankKeys.push(row._sourceRow);return;}
      const amount = parseAmount(row[amountColumn]);
      if(!groups.has(key)) groups.set(key,{key,displayKey:String(rawKey).trim(),amount:0,count:0,rows:[],source:sourceLabel});
      const group = groups.get(key);
      group.amount = roundMoney(group.amount + amount);
      group.count += 1;
      group.rows.push(row._sourceRow);
    });
    return {groups,blankKeys};
  }

  function reconcile(orderRows,paymentRows,config){
    const tolerance = Math.max(0,parseAmount(config.tolerance ?? 0.01));
    const order = groupRows(orderRows,config.orderKey,config.orderAmount,'订单表');
    const payment = groupRows(paymentRows,config.paymentKey,config.paymentAmount,'收款表');
    const keys = [...new Set([...order.groups.keys(),...payment.groups.keys()])];
    const results = keys.map(key=>{
      const left = order.groups.get(key);
      const right = payment.groups.get(key);
      const orderAmount = left?.amount ?? 0;
      const paymentAmount = right?.amount ?? 0;
      const difference = roundMoney(paymentAmount-orderAmount);
      const duplicate = (left?.count ?? 0)>1 || (right?.count ?? 0)>1;
      let primaryStatus;
      if(!right) primaryStatus = STATUS.MISSING;
      else if(!left) primaryStatus = STATUS.EXTRA;
      else if(Math.abs(difference)>tolerance) primaryStatus = STATUS.DIFFERENCE;
      else if(duplicate) primaryStatus = STATUS.DUPLICATE;
      else primaryStatus = STATUS.MATCHED;
      const flags=[];
      if((left?.count ?? 0)>1) flags.push('订单表重复');
      if((right?.count ?? 0)>1) flags.push('收款表重复');
      if(left&&right&&Math.abs(difference)>tolerance) flags.push('金额超出误差');
      return {
        status:primaryStatus,
        key:left?.displayKey || right?.displayKey || key,
        orderCount:left?.count ?? 0,
        orderAmount,
        paymentCount:right?.count ?? 0,
        paymentAmount,
        difference,
        orderRows:left?.rows.join('、') || '',
        paymentRows:right?.rows.join('、') || '',
        flags:flags.join('；')
      };
    });
    results.sort((a,b)=>{
      const rank={[STATUS.DIFFERENCE]:1,[STATUS.MISSING]:2,[STATUS.EXTRA]:3,[STATUS.DUPLICATE]:4,[STATUS.MATCHED]:5};
      return rank[a.status]-rank[b.status] || String(a.key).localeCompare(String(b.key),'zh-CN',{numeric:true});
    });
    const counts=Object.fromEntries(Object.values(STATUS).map(value=>[value,0]));
    results.forEach(row=>{counts[row.status]=(counts[row.status]||0)+1});
    const totalOrder = roundMoney([...order.groups.values()].reduce((sum,item)=>sum+item.amount,0));
    const totalPayment = roundMoney([...payment.groups.values()].reduce((sum,item)=>sum+item.amount,0));
    return {
      results,counts,tolerance,
      total:results.length,
      matched:counts[STATUS.MATCHED],
      issues:results.length-counts[STATUS.MATCHED],
      totalOrder,totalPayment,
      totalDifference:roundMoney(totalPayment-totalOrder),
      blankOrderKeys:order.blankKeys,
      blankPaymentKeys:payment.blankKeys
    };
  }

  function safeSpreadsheetValue(value){
    if(typeof value !== 'string') return value;
    return /^[=+\-@\t\r]/.test(value) ? `'${value}` : value;
  }

  global.ReconcileCore={MAX_ROWS,STATUS,normalizeKey,parseAmount,roundMoney,makeHeaders,rowsFromMatrix,reconcile,safeSpreadsheetValue};
})(window);

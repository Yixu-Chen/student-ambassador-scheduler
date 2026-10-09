(function(root,factory){const api=factory();if(typeof module==='object'&&module.exports)module.exports=api;else root.ScheduleEngine=api})(typeof self!=='undefined'?self:this,function(){
  const weekdayText=['周日','周一','周二','周三','周四','周五','周六'];
  function pad(n){return String(n).padStart(2,'0')}
  function dateKey(d){return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`}
  function groupKey(v=''){return v.endsWith('组组长')?v.slice(0,-2):v.endsWith('组长')?v.slice(0,-2):v}
  function buildShifts(settings){
    const out=[],days=new Date(settings.year,settings.month,0).getDate();
    for(let n=1;n<=days;n++){
      const d=new Date(settings.year,settings.month-1,n),w=d.getDay();
      if(!((w===6&&settings.sat)||(w===0&&settings.sun)))continue;
      const date=dateKey(d),prefix=`${settings.month}月${pad(n)}日${weekdayText[w]}`;
      out.push({id:prefix+'上午',date,weekday:weekdayText[w],period:'上午',start:settings.amStart,end:settings.amEnd,required:settings.headcount});
      out.push({id:prefix+'下午',date,weekday:weekdayText[w],period:'下午',start:settings.pmStart,end:settings.pmEnd,required:settings.headcount});
    }
    return out;
  }
  function allowedByRule(person,shift){
    const r=person.rule||'均可';
    if(r==='均可')return true;
    if(r==='仅周六下午')return shift.weekday==='周六'&&shift.period==='下午';
    if(r==='仅周六上午')return shift.weekday==='周六'&&shift.period==='上午';
    if(r==='仅周日下午')return shift.weekday==='周日'&&shift.period==='下午';
    if(r==='仅周日上午')return shift.weekday==='周日'&&shift.period==='上午';
    return true;
  }
  function isBlocked(name,shift,blocks){return blocks.some(b=>b.name===name&&b.date===shift.date&&(b.period==='全天'||b.period===shift.period))}
  function eligible(person,shift,counts,blocks){return counts[person.name]<(Number(person.max)||0)&&allowedByRule(person,shift)&&!isBlocked(person.name,shift,blocks)}
  function teamScore(team,counts,opts){
    let score=0;
    for(const p of team){if(counts[p.name]<(Number(p.min)||0))score+=10000;if(opts.coverage&&counts[p.name]===0)score+=1000;score-=10*(Number(p.history)||0)}
    const types=new Set(team.map(p=>p.type));
    if(opts.mixed&&types.has('老人')&&types.has('新人'))score+=500;
    if(opts.group&&types.has('老人')&&types.has('新人')){
      const olds=new Set(team.filter(p=>p.type==='老人').map(p=>groupKey(p.group)));
      if(team.some(p=>p.type==='新人'&&olds.has(groupKey(p.group))))score+=100;
    }
    return score;
  }
  function combinations(list,k,limit=50000){const out=[];function go(i,a){if(out.length>=limit)return;if(a.length===k){out.push(a.slice());return}for(let j=i;j<=list.length-(k-a.length);j++){a.push(list[j]);go(j+1,a);a.pop()}}go(0,[]);return out}
  function schedule(input){
    const {people,blocks=[],settings}=input,shifts=buildShifts(settings),counts=Object.fromEntries(people.map(p=>[p.name,0])),assignments={};
    if(!people.length)throw new Error('请先上传或添加人员');
    for(const shift of shifts){
      const candidates=people.filter(p=>eligible(p,shift,counts,blocks));
      if(candidates.length<shift.required)throw new Error(`${shift.id}可排人员不足：需要${shift.required}人，只有${candidates.length}人`);
      let teams=combinations(candidates,shift.required);
      if(!teams.length||teams.length>=50000){teams=[candidates.slice().sort((a,b)=>(counts[a.name]+a.history)-(counts[b.name]+b.history)).slice(0,shift.required)]}
      teams.sort((a,b)=>teamScore(b,counts,settings)-teamScore(a,counts,settings)||a.map(x=>x.name).join('').localeCompare(b.map(x=>x.name).join(''),'zh-CN'));
      assignments[shift.id]=teams[0].map(p=>p.name);for(const p of teams[0])counts[p.name]++;
    }
    const unmet=people.filter(p=>counts[p.name]<(Number(p.min)||0));
    if(unmet.length)throw new Error(`未满足最低次数：${unmet.map(p=>p.name).join('、')}`);
    return {shifts,assignments,counts};
  }
  function validate(result,people,blocks,settings){
    const warnings=[],byName=Object.fromEntries(people.map(p=>[p.name,p]));const counts=Object.fromEntries(people.map(p=>[p.name,0]));let mixed=0,sameGroup=0;
    for(const s of result.shifts){const names=result.assignments[s.id]||[];if(names.length!==s.required)warnings.push(`${s.id}人数应为${s.required}人`);if(new Set(names).size!==names.length)warnings.push(`${s.id}存在重复人员`);const team=names.map(n=>byName[n]).filter(Boolean);for(const p of team){counts[p.name]++;if(!allowedByRule(p,s)||isBlocked(p.name,s,blocks))warnings.push(`${p.name}不适合${s.id}`)}const types=new Set(team.map(p=>p.type));if(types.has('老人')&&types.has('新人')){mixed++;const oldGroups=new Set(team.filter(p=>p.type==='老人').map(p=>groupKey(p.group)));if(team.some(p=>p.type==='新人'&&oldGroups.has(groupKey(p.group))))sameGroup++}}
    for(const p of people){if(counts[p.name]<(Number(p.min)||0))warnings.push(`${p.name}未达到最低次数`);if(counts[p.name]>(Number(p.max)||0))warnings.push(`${p.name}超过最多次数`)}
    return {warnings,counts,unique:Object.values(counts).filter(Boolean).length,total:Object.values(counts).reduce((a,b)=>a+b,0),mixed,sameGroup};
  }
  return {buildShifts,schedule,validate,groupKey,dateKey};
});

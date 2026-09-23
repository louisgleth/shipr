(() => {
  const byId=id=>document.getElementById(id),section=byId('adminPageSection'),dialog=byId('adminAnalyticsDialog'),frame=byId('adminAnalyticsFrame');
  if(!section||!dialog)return;
  let busy=false,summaryAt=0,live=null,generation=0,reportGeneration=0,wasActive=false;
  const active=()=>Boolean(currentUser&&adminAccessAllowed&&!section.classList.contains('is-hidden'));
  const text=(id,value)=>{byId(id).textContent=value;};
  function theme(){
    const doc=frame.contentDocument;if(!doc)return;
    const style=getComputedStyle(document.documentElement);
    for(const key of ['--primary','--secondary','--surface','--surface-2','--text','--muted','--stroke','--accent','--radius','--radius-lg'])doc.documentElement.style.setProperty(key,style.getPropertyValue(key));
  }
  function renderLive(){
    if(!live||!dialog.open)return;const doc=frame.contentDocument;if(!doc?.getElementById('live-count'))return;
    doc.getElementById('live-count').textContent=live.total;
    doc.getElementById('live-status').textContent='Updated '+new Date(live.generatedAt).toLocaleTimeString()+' · Refreshes every 10 seconds';
    const sites=doc.getElementById('live-sites'),rows=doc.getElementById('live-rows');sites.replaceChildren();rows.replaceChildren();
    for(const site of live.sites){const card=doc.createElement('div');card.className='live-site';const label=doc.createElement('span'),n=doc.createElement('strong');label.textContent=site.name;n.textContent=site.active;card.append(label,n);sites.append(card);}
    for(const v of live.visitors){const row=doc.createElement('tr');for(const value of [live.sites.find(s=>s.id===v.site)?.name||v.site,v.path,v.device,v.locale,v.source||v.referrer||'Direct / unknown',Math.max(0,Math.round((Date.parse(live.generatedAt)-Date.parse(v.lastSeen))/1000))+'s ago']){const td=doc.createElement('td');td.textContent=value??'—';row.append(td);}rows.append(row);}
    if(!live.visitors.length){const row=doc.createElement('tr'),td=doc.createElement('td');td.colSpan=6;td.textContent='No visitors active right now.';row.append(td);rows.append(row);}
  }
  function clear(){generation++;reportGeneration++;summaryAt=0;live=null;if(dialog.open)dialog.close();frame.removeAttribute('srcdoc');for(const id of ['Live','Sessions','Clicks','Leads','ArticleViews','ArticleGoogleClicks'])text('adminAnalytics'+id,'—');}
  async function refresh(){
    if(!active()){if(wasActive)clear();wasActive=false;return;}wasActive=true;
    if(busy||document.visibilityState==='hidden')return;busy=true;const ticket=generation,userId=currentUser.id;
    try{
      if(Date.now()-summaryAt>300000){const data=await fetchApiWithAuth('/api/admin/analytics/summary');if(ticket!==generation||!active()||currentUser.id!==userId)return;
        live=data.live;summaryAt=Date.now();text('adminAnalyticsSessions',data.sessions.toLocaleString());text('adminAnalyticsClicks',data.clicks.toLocaleString());text('adminAnalyticsLeads',data.leads.toLocaleString());
        text('adminAnalyticsArticleViews',data.articleViews?.toLocaleString()??'—');text('adminAnalyticsArticleGoogleClicks',data.articleGoogleClicks?.toLocaleString()??'—');
      }else{const data=await fetchApiWithAuth('/api/admin/analytics/live');if(ticket!==generation||!active()||currentUser.id!==userId)return;live=data;}
      text('adminAnalyticsLive',live.total.toLocaleString());text('adminAnalyticsStatus','Live · Updated '+new Date(live.generatedAt).toLocaleTimeString());renderLive();
    }catch(error){if(ticket!==generation)return;text('adminAnalyticsLive','—');text('adminAnalyticsStatus','Analytics unavailable · retrying shortly');const doc=frame.contentDocument;if(doc?.getElementById('live-count')){doc.getElementById('live-count').textContent='—';doc.getElementById('live-status').textContent='Live updates unavailable · retrying';doc.getElementById('live-sites').replaceChildren();doc.getElementById('live-rows').replaceChildren();}}
    finally{busy=false;}
  }
  async function loadReport(){
    if(!active()||!dialog.open)return;text('adminAnalyticsReportStatus','Loading report…');byId('adminAnalyticsReportStatus').hidden=false;frame.hidden=true;const ticket=generation,reportTicket=++reportGeneration,userId=currentUser.id;
    try{const {html}=await fetchApiWithAuth('/api/admin/analytics/report?days='+byId('adminAnalyticsPeriod').value,{timeoutMs:20000});if(ticket!==generation||reportTicket!==reportGeneration||!active()||currentUser.id!==userId||!dialog.open)return;
      const doc=new DOMParser().parseFromString(html,'text/html');doc.querySelectorAll('script,style').forEach(el=>el.remove());
      const css=doc.createElement('link');css.rel='stylesheet';css.href=new URL('/admin-analytics-report.css?v=20260923',location.origin).href;doc.head.append(css);
      frame.srcdoc='<!doctype html>'+doc.documentElement.outerHTML;
    }catch{if(ticket===generation&&reportTicket===reportGeneration)text('adminAnalyticsReportStatus','Could not load analytics. Close this window and try again.');}
  }
  byId('openAdminAnalytics').addEventListener('click',()=>{if(!active())return;dialog.showModal();void loadReport();});
  byId('adminAnalyticsPeriod').addEventListener('change',()=>{void loadReport();});
  frame.addEventListener('load',()=>{
    if(!dialog.open||!frame.getAttribute('srcdoc'))return;
    const doc=frame.contentDocument;
    // Fragment links in srcdoc otherwise navigate to the parent portal URL.
    doc?.addEventListener('click',event=>{
      const link=event.target.closest('a[href^="#"]');
      if(!link)return;event.preventDefault();doc.getElementById(link.getAttribute('href').slice(1))?.scrollIntoView({block:'start'});
    });
    theme();renderLive();frame.hidden=false;byId('adminAnalyticsReportStatus').hidden=true;
  });
  byId('closeAdminAnalytics').addEventListener('click',()=>dialog.close());
  dialog.addEventListener('close',()=>{reportGeneration++;frame.removeAttribute('srcdoc');byId('openAdminAnalytics').focus();});
  dialog.addEventListener('click',e=>{if(e.target===dialog){const r=dialog.getBoundingClientRect();if(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom)dialog.close();}});
  new MutationObserver(refresh).observe(section,{attributes:true,attributeFilter:['class']});
  new MutationObserver(theme).observe(document.documentElement,{attributes:true,attributeFilter:['data-theme']});
  document.addEventListener('visibilitychange',refresh);setInterval(refresh,10000);refresh();
})();

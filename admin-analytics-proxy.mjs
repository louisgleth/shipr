const response=(status,data)=>Response.json(data,{status,headers:{'Cache-Control':'no-store','X-Content-Type-Options':'nosniff'}});
export async function handleAdminAnalytics(request,env,{authenticate,isAdmin}) {
 const user=await authenticate(request,env);
 if(!user?.id)return response(401,{error:'Authentication required.'});
 if(!isAdmin(user,env))return response(403,{error:'Admin access required.'});
 if(!env.OBSERVATORY||!env.OBSERVATORY_ADMIN_TOKEN)return response(503,{error:'Analytics is not connected yet.'});
 const view=new URL(request.url).pathname.split('/').pop();
 if(!['summary','live','report'].includes(view))return response(404,{error:'Analytics view not found.'});
 async function read(path,json=true){
  const result=await env.OBSERVATORY.fetch(new Request('https://observatory.internal'+path,{headers:{Authorization:'Bearer '+env.OBSERVATORY_ADMIN_TOKEN},signal:AbortSignal.timeout(15000)}));
  if(!result.ok)throw new Error('Analytics unavailable');
  return json?result.json():result.text();
 }
 try{
  if(view==='live')return response(200,await read('/v1/live'));
  if(view==='report')return response(200,{html:await read('/?days=28',false)});
  const [report,live]=await Promise.all([read('/v1/report?days=28'),read('/v1/live')]);
  const tools=report.sites.filter(s=>s.id!=='shipide');
  return response(200,{live,days:28,sessions:tools.reduce((n,s)=>n+s.sessions,0),clicks:tools.reduce((n,s)=>n+(s.counts.shipide_click||0),0),leads:report.sites.find(s=>s.id==='shipide')?.leads||0,googleClicks:tools.some(s=>s.search.clicks!==null)?tools.reduce((n,s)=>n+(s.search.clicks||0),0):null});
 }catch{return response(502,{error:'Analytics is temporarily unavailable. Please try again.'});}
}

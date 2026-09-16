import {test} from 'node:test';
import assert from 'node:assert/strict';
import {handleAdminAnalytics} from '../admin-analytics-proxy.mjs';
const request=view=>new Request('https://portal.shipide.com/api/admin/analytics/'+view);
test('analytics never reaches collector for unsigned or non-admin users',async()=>{
 let calls=0;const env={OBSERVATORY:{fetch(){calls++;throw Error();}},OBSERVATORY_ADMIN_TOKEN:'secret'};
 assert.equal((await handleAdminAnalytics(request('live'),env,{authenticate:async()=>null,isAdmin:()=>true})).status,401);
 assert.equal((await handleAdminAnalytics(request('live'),env,{authenticate:async()=>({id:'client',user_metadata:{app_admin:true}}),isAdmin:()=>false})).status,403);
 assert.equal(calls,0);
});
test('admin proxy limits destinations, uses server credential, and avoids duplicate leads',async()=>{
 const calls=[];const env={OBSERVATORY_ADMIN_TOKEN:'server-secret',OBSERVATORY:{fetch:async req=>{calls.push(req);return Response.json(req.url.endsWith('/v1/live')?{total:2}:{sites:[{id:'marge',sessions:3,counts:{shipide_click:2},leads:1,search:{clicks:4}},{id:'shipide',sessions:1,counts:{},leads:1,search:{clicks:1}}]});}}};
 const auth={authenticate:async()=>({id:'admin'}),isAdmin:()=>true};
 const result=await handleAdminAnalytics(request('summary'),env,auth),body=await result.json();assert.equal(body.leads,1);assert.equal(body.sessions,3);assert.equal(body.live.total,2);assert.equal(JSON.stringify(body).includes('server-secret'),false);assert.ok(calls.every(r=>r.headers.get('Authorization')==='Bearer server-secret'));assert.equal(result.headers.get('Cache-Control'),'no-store');
 assert.equal((await handleAdminAnalytics(request('arbitrary'),env,auth)).status,404);assert.equal(calls.length,2);
 env.OBSERVATORY.fetch=async()=>new Response('upstream secret error',{status:500});const failure=await handleAdminAnalytics(request('live'),env,auth);assert.equal(failure.status,502);assert.equal((await failure.text()).includes('upstream secret'),false);
});

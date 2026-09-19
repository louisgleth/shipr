import { writeFileSync } from 'node:fs';

const ref = name => ({ $ref: `#/components/schemas/${name}` });
const string = { type:'string' };
const object = (properties, required = Object.keys(properties), extra = {}) => ({ type:'object', properties, required, ...extra });
const envelope = schema => object({ data:schema });
const id = { type:'string', description:'Opaque Shipide resource ID.' };
const date = { type:'string', format:'date-time' };
const address = object(Object.fromEntries(['name','company','street1','street2','city','state','postal_code','country','email','phone'].map(key => [key, { type:'string', minLength:key === 'country' ? 2 : 1, maxLength:key === 'country' ? 2 : 200 }])), ['name','street1','city','postal_code','country'], { additionalProperties:false });
const parcel = object(Object.fromEntries(['weight_kg','length_cm','width_cm','height_cm'].map(key => [key, { type:'number', exclusiveMinimum:0, maximum:1000 }])), undefined, { additionalProperties:false });
const input = object({ reference:{ type:'string', minLength:1, maxLength:200 }, from:ref('Address'), to:ref('Address'), parcel:ref('Parcel'), order_id:id, metadata:{ type:'object', additionalProperties:true, description:'Maximum 4096 characters when serialized.' } }, ['reference','from','to','parcel'], { additionalProperties:false });
const orderInput = structuredClone(input); delete orderInput.properties.order_id;
const shipment = object({ ...input.properties, id, object:{ const:'shipment' }, status:{ enum:['draft','label_created','voided'] }, livemode:{ type:'boolean' }, created_at:date, label_id:id, tracking_number:string }, ['id','object','reference','from','to','parcel','status','livemode','created_at']);
const order = object({ ...orderInput.properties, id, object:{ const:'order' }, status:{ const:'received' }, livemode:{ type:'boolean' }, created_at:date }, ['id','object','reference','from','to','parcel','status','livemode','created_at']);
const rate = object({ id, shipment_id:id, carrier:{ const:'shipide_sandbox' }, service:{ const:'sandbox_standard' }, currency:{ const:'EUR' }, amount:{ type:'number', const:5 }, test:{ const:true } });
const label = object({ id, object:{ const:'label' }, shipment_id:id, status:{ enum:['created','voided'] }, carrier:{ const:'shipide_sandbox' }, tracking_number:string, format:{ const:'pdf' }, download_path:string, charge:object({ currency:{ const:'EUR' }, amount:{ type:'number', const:0 } }), test:{ const:true }, created_at:date, voided_at:date }, ['id','object','shipment_id','status','carrier','tracking_number','format','download_path','charge','test','created_at']);
const webhook = object({ id, url:{ type:'string', format:'uri' }, events:{ type:'array', items:{ enum:['order.created','shipment.created','label.created','label.voided'] }, minItems:1, maxItems:4 }, created_at:date, disabled_at:{ type:['string','null'], format:'date-time' } }, ['id','url','events','created_at']);
const keyParam = { name:'Idempotency-Key', in:'header', required:true, description:'Persist a unique key per logical operation. Retry the identical request with the same key. Successful keys do not currently expire.', schema:{ type:'string', minLength:1, maxLength:128 } };
const idParam = name => ({ name, in:'path', required:true, schema:id });
const pagination = [{ name:'limit', in:'query', schema:{ type:'integer', minimum:1, maximum:100, default:25 } }, { name:'after', in:'query', description:'Use next_cursor from the previous page. ID order, not chronological order.', schema:string }];
const response = schema => ({ description:'Success', headers:{ 'X-Request-Id':{ schema:string }, 'Idempotency-Replayed':{ schema:{ type:'string', enum:['true'] }, description:'Present when returning a stored response.' } }, content:{ 'application/json':{ schema } } });
const errorResponses = Object.fromEntries([400,401,403,404,409,413,415,422,429,500,503].map(status => [String(status), { description:`${status} API error; inspect error.code and error.message.`, content:{ 'application/json':{ schema:ref('Error') } } }]));
const paths = {};
function operation(path, method, operationId, scope, summary, schema, options = {}) {
  const parameters = [...(options.parameters || []), ...(['post','delete'].includes(method) ? [keyParam] : [])];
  const op = { operationId, summary, description:`Requires the ${scope} scope.${options.description ? ` ${options.description}` : ''}`, tags:[options.tag || path.split('/')[1]], security:[{ ApiKey:[] }], parameters, responses:{ ...errorResponses, [options.status || 200]:response(schema) } };
  if (options.body) op.requestBody = { required:true, content:{ 'application/json':{ schema:options.body } } };
  (paths[path] ||= {})[method] = op;
  return op;
}
operation('/account','get','getAccount','account:read','Get account and capabilities', envelope(ref('Account')));
for (const [collection, model] of [['orders','Order'],['shipments','Shipment'],['labels','Label']]) {
  operation(`/${collection}`,'get',`list${model}s`,`${collection}:read`,`List ${collection}`, object({ data:{ type:'array', items:ref(model) }, next_cursor:{ type:['string','null'] } }), { parameters:pagination });
  operation(`/${collection}/{id}`,'get',`get${model}`,`${collection}:read`,`Retrieve ${model.toLowerCase()}`, envelope(ref(model)), { parameters:[idParam('id')] });
  if (collection !== 'labels') operation(`/${collection}`,'post',`create${model}`,`${collection}:write`,`Create ${model.toLowerCase()}`, envelope(ref(model)), { status:201, body:ref(`${model}Input`) });
}
operation('/rates','post','getRates','rates:read','Get sandbox rates', envelope({ type:'array', items:ref('Rate') }), { body:object({ shipment_id:id }, undefined, { additionalProperties:false }), description:'Sandbox only. Live requests return 503 carrier_not_configured without charge.' });
operation('/labels','post','createLabel','labels:write','Create a sandbox label', envelope(ref('Label')), { status:201, body:object({ shipment_id:id, rate_id:id, format:{ type:'string', enum:['pdf'], default:'pdf' } }, ['shipment_id','rate_id'], { additionalProperties:false }), description:'Sandbox only. Test PDFs are not valid postage. No charges are made.' });
operation('/labels/{id}/void','post','voidLabel','labels:write','Void a sandbox label', envelope(ref('Label')), { parameters:[idParam('id')], body:object({}, [], { additionalProperties:false }) });
const file = operation('/labels/{id}/file','get','downloadLabel','labels:read','Download a sandbox PDF', {}, { parameters:[idParam('id')] });
file.responses['200'] = { description:'4 by 6 inch test label. Not valid for shipping.', content:{ 'application/pdf':{ schema:{ type:'string', format:'binary' } } } };
operation('/tracking/{tracking_number}','get','getTracking','shipments:read','Retrieve sandbox tracking', envelope(object({ tracking_number:string, shipment_id:id, status:{ enum:['pre_transit','voided'] }, test:{ const:true }, events:{ type:'array', maxItems:0, items:{ type:'object' } } })), { parameters:[idParam('tracking_number')] });
operation('/webhooks','post','createWebhook','webhooks:write','Create a webhook subscription', envelope(object({ ...webhook.properties, secret:{ type:'string', description:'Signing secret returned on create and identical idempotent replays only.' } }, ['id','url','events','created_at','secret'])), { status:201, body:object({ url:{ type:'string', format:'uri', maxLength:2000 }, events:webhook.properties.events }, undefined, { additionalProperties:false }) });
operation('/webhooks','get','listWebhooks','webhooks:read','List webhook subscriptions', envelope({ type:'array', items:ref('Webhook') }));
operation('/webhooks/{id}','get','getWebhook','webhooks:read','Retrieve webhook subscription', envelope(ref('Webhook')), { parameters:[idParam('id')] });
operation('/webhooks/{id}','delete','disableWebhook','webhooks:write','Disable webhook subscription', envelope(object({ id, disabled:{ const:true } })), { parameters:[idParam('id')] });
const spec = {
  openapi:'3.1.0',
  info:{ title:'Shipide API', version:'1.0.0', description:'Integrate order and warehouse systems with Shipide. Live order and shipment intake; sandbox rates, labels, tracking and voids. Live carrier booking is not yet configured. Sandbox operations never charge your account.' },
  servers:[{ url:'https://portal.shipide.com/api/v1' }],
  security:[{ ApiKey:[] }], paths,
  components:{ securitySchemes:{ ApiKey:{ type:'http', scheme:'bearer', bearerFormat:'shipide_test_... or shipide_live_...', description:'Create scoped API keys in Portal > Account > Developer.' } }, schemas:{
    Address:address, Parcel:parcel, ShipmentInput:input, OrderInput:orderInput, Shipment:shipment, Order:order, Rate:rate, Label:label, Webhook:webhook,
    Error:object({ error:object({ code:string, message:string, request_id:string }) }),
    Account:object({ id, mode:{ enum:['test','live'] }, scopes:{ type:'array', items:string }, capabilities:object({ orders:{ const:true }, shipments:{ const:true }, labels:{ type:'boolean' }, rates:{ type:'boolean' }, tracking:{ const:'sandbox_only' }, live_carrier_booking:{ const:false } }), docs_url:{ type:'string', format:'uri' } }),
  } },
};
writeFileSync(new URL('../docs/openapi.json', import.meta.url), JSON.stringify(spec, null, 2) + '\n');
console.log(`Generated OpenAPI 3.1: ${Object.keys(paths).length} paths, ${Object.values(paths).reduce((n, p) => n + Object.keys(p).length, 0)} operations.`);

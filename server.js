import express from 'express';
import cors from 'cors';
import helmet from 'helmet';
import multer from 'multer';
import Database from 'better-sqlite3';
import crypto from 'crypto';
import dotenv from 'dotenv';
import OpenAI from 'openai';
import { Document, Packer, Paragraph, HeadingLevel } from 'docx';
import pptxgen from 'pptxgenjs';
import fs from 'fs';
import path from 'path';

dotenv.config();
const app=express();
const PORT=Number(process.env.PORT||3000);
const TIMEOUT=Number(process.env.SESSION_TIMEOUT_MINUTES||30)*60*1000;
const MAX_MB=Number(process.env.MAX_FILE_MB||10);
const db=new Database('data.sqlite');
db.pragma('journal_mode = WAL');
for(const s of fs.readFileSync('schema.sql','utf8').split(';')) if(s.trim()) db.prepare(s).run();
const upload=multer({dest:'uploads/',limits:{fileSize:MAX_MB*1024*1024},fileFilter:(req,file,cb)=>{const ok=['image/png','image/jpeg','application/pdf'].includes(file.mimetype);cb(ok?null:new Error('Yalnız PNG, JPG/JPEG və PDF fayllarına icazə verilir.'),ok)}});
app.use(helmet({crossOriginResourcePolicy:false}));
app.use(cors()); app.use(express.json({limit:'1mb'})); app.use(express.static('public'));
const now=()=>new Date().toISOString();
const hash=x=>crypto.createHash('sha256').update(x).digest('hex');
function session(req,res,next){const sid=req.header('x-session-id');if(!sid)return res.status(401).json({error:'Sessiya tələb olunur'});const s=db.prepare('SELECT * FROM sessions WHERE id=? AND active=1').get(sid);if(!s||Date.now()-Date.parse(s.last_activity)>TIMEOUT){if(s)closeSession(s);return res.status(401).json({error:'Sessiya bitib. Yenidən daxil olun.'})}db.prepare('UPDATE sessions SET last_activity=?,expires_at=? WHERE id=?').run(now(),new Date(Date.now()+TIMEOUT).toISOString(),sid);db.prepare('UPDATE access_codes SET last_activity=? WHERE id=?').run(now(),s.access_code_id);req.session=s;next()}
function closeSession(s){db.prepare('UPDATE sessions SET active=0 WHERE id=?').run(s.id);db.prepare('UPDATE access_codes SET active_session_id=NULL,logout_time=? WHERE id=?').run(now(),s.access_code_id)}
app.post('/api/login',(req,res)=>{const code=String(req.body.code||'').trim();if(!code)return res.status(400).json({error:'Giriş kodunu daxil edin'});const c=db.prepare('SELECT * FROM access_codes WHERE code_hash=?').get(hash(code));if(!c||!c.active|| (c.expires_at&&Date.parse(c.expires_at)<Date.now()))return res.status(403).json({error:'Yanlış, deaktiv və ya müddəti bitmiş giriş kodu'});if(c.active_session_id){const s=db.prepare('SELECT * FROM sessions WHERE id=? AND active=1').get(c.active_session_id);if(s&&Date.now()-Date.parse(s.last_activity)<=TIMEOUT)return res.status(409).json({error:'Bu giriş kodu hazırda başqa aktiv sessiyada istifadə olunur. Digər sessiya bağlandıqdan sonra yenidən cəhd edin.'});if(s)closeSession(s)}const sid=crypto.randomUUID();const t=now();db.prepare('INSERT INTO sessions(id,access_code_id,created_at,last_activity,expires_at,device_info) VALUES(?,?,?,?,?,?)').run(sid,c.id,t,t,new Date(Date.now()+TIMEOUT).toISOString(),req.get('user-agent')||'');db.prepare('UPDATE access_codes SET active_session_id=?,login_time=?,last_activity=? WHERE id=?').run(sid,t,t,c.id);res.json({session_id:sid,timeout_minutes:TIMEOUT/60000})});
app.post('/api/logout',session,(req,res)=>{closeSession(req.session);res.json({ok:true})});
app.get('/api/dashboard',session,(req,res)=>{const count=db.prepare('SELECT COUNT(*) n FROM materials WHERE session_id=?').get(req.session.id).n;const recent=db.prepare('SELECT id,title,module,updated_at FROM materials WHERE session_id=? ORDER BY updated_at DESC LIMIT 8').all(req.session.id);res.json({date:new Date().toLocaleDateString('az-AZ'),count,recent})});
const moduleNames={
 'teqvim-tematik':'Təqvim-tematik plan','ders-plani':'Dərs planı','ders-ssenarisi':'Dərs ssenarisi','interaktiv-metodlar':'İnteraktiv metodlar','ksq':'Kiçik summativ qiymətləndirmə','bsq':'Böyük summativ qiymətləndirmə','telim-neticesi':'Təlim nəticələrinin qiymətləndirilməsi','seristte':'Səriştə qiymətləndirilməsi','netice':'Nəticələrin yoxlanılması','ev-tapsirigi':'Ev tapşırığı','test':'Test və sual generatoru','ppt':'PowerPoint təqdimatı','chat':'AI Çat','qanun':'Təhsil qanunvericiliyi','kurikulum':'Kurikulum və standartlar','fayl':'Fayl analizi'};
const system=`Sən Müəllimin AI Köməkçisisən. Azərbaycan təhsil sistemi üçün peşəkar müəllim, metodist, kurikulum, qiymətləndirmə və peşə təhsili mütəxəssisi kimi cavab ver. Cavab Azərbaycan dilində, strukturlaşdırılmış və praktik olsun. Hüquqi/metodik iddiaları uydurma; təsdiqlənməyən hüquqi məlumat üçün “Bu məlumat üzrə etibarlı rəsmi mənbə müəyyən edilmədi.” yaz. Mənbə verilirsə sənəd adı, maddə/bənd və rəsmi URL göstər.`;
let client=null; if(process.env.OPENAI_API_KEY) client=new OpenAI({apiKey:process.env.OPENAI_API_KEY});
async function ai(prompt){if(!client)return 'AI xidməti üçün OPENAI_API_KEY server mühitində konfiqurasiya edilməlidir.';const r=await client.responses.create({model:'gpt-5-mini',input:[{role:'system',content:system},{role:'user',content:prompt}]});return r.output_text||'Nəticə alınmadı.'}
app.post('/api/ai',session,upload.single('file'),async(req,res)=>{try{const module=moduleNames[req.body.module]||req.body.module||'AI Çat';const prompt=`Modul: ${module}\nİstək: ${req.body.prompt||''}\nParametrlər: ${req.body.params||''}${req.file?'\nFayl yüklənib: '+req.file.originalname:' '}\nMümkün qədər birbaşa istifadə edilə bilən nəticə hazırla.`;const content=await ai(prompt);const title=req.body.title||module;db.prepare('INSERT INTO materials(session_id,title,module,content,created_at,updated_at) VALUES(?,?,?,?,?,?)').run(req.session.id,title,module,content,now(),now());res.json({title,module,content})}catch(e){res.status(500).json({error:e.message})}});
app.get('/api/materials',session,(req,res)=>res.json(db.prepare('SELECT * FROM materials WHERE session_id=? ORDER BY updated_at DESC').all(req.session.id)));
app.put('/api/materials/:id',session,(req,res)=>{db.prepare('UPDATE materials SET title=?,content=?,updated_at=? WHERE id=? AND session_id=?').run(req.body.title,req.body.content,now(),req.params.id,req.session.id);res.json({ok:true})});
app.delete('/api/materials/:id',session,(req,res)=>{db.prepare('DELETE FROM materials WHERE id=? AND session_id=?').run(req.params.id,req.session.id);res.json({ok:true})});
app.post('/api/export/docx',session,async(req,res)=>{const lines=String(req.body.content||'').split('\n');const children=[new Paragraph({text:req.body.title||'Müəllimin AI Köməkçisi',heading:HeadingLevel.TITLE}),...lines.map(x=>new Paragraph(x))];const doc=new Document({sections:[{children}]});const buf=await Packer.toBuffer(doc);res.set({'Content-Type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document','Content-Disposition':'attachment; filename="muellimin-ai-komekcisi.docx"'});res.send(buf)});
app.post('/api/export/pptx',session,async(req,res)=>{const ppt=new pptxgen();ppt.layout='LAYOUT_WIDE';const slides=Array.isArray(req.body.slides)?req.body.slides:[];for(const x of slides){const s=ppt.addSlide();s.addText(x.title||'', {x:.6,y:.5,w:12,h:.6,fontSize:26,bold:true});s.addText(x.body||'', {x:.7,y:1.4,w:11.8,h:5.2,fontSize:18,breakLine:false,fit:'shrink'});}const b=await ppt.write({outputType:'nodebuffer'});res.set({'Content-Type':'application/vnd.openxmlformats-officedocument.presentationml.presentation','Content-Disposition':'attachment; filename="muellimin-ai-komekcisi.pptx"'});res.send(b)});
function admin(req,res,next){const a=req.headers.authorization||'';const [u,p]=Buffer.from(a.replace('Basic ',''),'base64').toString().split(':');if(u!==process.env.ADMIN_USERNAME||p!==process.env.ADMIN_PASSWORD)return res.status(401).json({error:'Admin giriş tələb olunur'});next()}
app.post('/api/admin/codes',admin,(req,res)=>{const raw=req.body.code||crypto.randomBytes(5).toString('hex').toUpperCase();const expires=req.body.expires_at||null;db.prepare('INSERT INTO access_codes(code_hash,label,expires_at) VALUES(?,?,?)').run(hash(raw),req.body.label||'Müəllim',expires);res.json({code:raw,expires_at:expires})});
app.get('/api/admin/sessions',admin,(req,res)=>res.json(db.prepare('SELECT s.id,c.label,s.created_at,s.last_activity,s.expires_at,s.device_info FROM sessions s JOIN access_codes c ON c.id=s.access_code_id WHERE s.active=1').all()));
app.post('/api/admin/sessions/:id/close',admin,(req,res)=>{const s=db.prepare('SELECT * FROM sessions WHERE id=?').get(req.params.id);if(s)closeSession(s);res.json({ok:true})});
app.get('/api/admin/codes',admin,(req,res)=>res.json(db.prepare('SELECT id,label,active,expires_at,active_session_id,login_time,last_activity,logout_time FROM access_codes ORDER BY id DESC').all()));
app.post('/api/admin/codes/:id/toggle',admin,(req,res)=>{db.prepare('UPDATE access_codes SET active=? WHERE id=?').run(req.body.active?1:0,req.params.id);res.json({ok:true})});
app.get('*',(req,res)=>res.sendFile(path.resolve('public/index.html')));
app.listen(PORT,()=>{fs.mkdirSync('uploads',{recursive:true});console.log(`Müəllimin AI Köməkçisi: http://localhost:${PORT}`)});

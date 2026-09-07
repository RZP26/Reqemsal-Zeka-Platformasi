# Müəllimin AI Köməkçisi

Bu branch statik HTML maketindən full-stack tətbiq arxitekturasına keçid üçün hazırlanıb.

## Stack
- Node.js + Express
- SQLite + better-sqlite3
- OpenAI Responses API
- DOCX: `docx`
- PPTX: `pptxgenjs`
- Multer + MIME/ölçü yoxlaması
- Helmet
- Responsive vanilla frontend

## Quraşdırma
1. Node.js 20+ quraşdırın.
2. `npm install`
3. `.env.example` faylını `.env` kimi kopyalayın.
4. `OPENAI_API_KEY`, `ADMIN_USERNAME`, `ADMIN_PASSWORD` dəyərlərini serverdə təyin edin. API açarı frontend-ə qoyulmur.
5. `npm start`
6. Müəllim interfeysi: `/`
7. Admin: `/admin.html`

## Təhlükəsizlik
- Giriş kodları verilənlər bazasında SHA-256 hash kimi saxlanır.
- Bir kod üçün eyni anda yalnız bir aktiv sessiya qəbul edilir.
- Fəaliyyətsizlik timeout-u server tərəfindən yoxlanılır.
- Admin aktiv sessiyanı bağlaya bilir.
- Yükləmə MIME type və ölçü ilə məhdudlaşdırılır.
- Real `.docx` və `.pptx` binary faylları serverdə yaradılır.
- Production mühitində HTTPS, reverse proxy, rate limiting, secret manager və müntəzəm backup əlavə edilməlidir.

## Vacib qeyd
Hüquqi və normativ cavabların real rəsmi mənbə ilə təsdiqi üçün admin panelində rəsmi mənbə bazası doldurulmalıdır. Sistem təsdiqlənməyən məlumatı hüquqi fakt kimi təqdim etməməlidir.

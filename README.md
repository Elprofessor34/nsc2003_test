# National School And College — Admin Portal

Cloud-synced student management web app for National School And College.

## Features

- Multi-admin login (each staff member has their own account)
- Real-time sync across devices — change made on one device appears on all others instantly
- 500+ students with photos
- Track monthly tuition, session fee, and 3 term exam fees
- Science / Humanities sections for Class Nine and Ten
- Payment history shows who recorded each payment and when
- Year-end class promotion tool
- CSV exports for Google Sheets

## How data is stored

All data lives in a Supabase cloud database. Sign in with email and password from any device — your data is always there.

## Local development

```
npm install
npm run dev
```

## Production build

```
npm run build
```

Deployed via Vercel — pushes to main branch auto-deploy.

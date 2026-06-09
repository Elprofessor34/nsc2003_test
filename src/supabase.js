import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL = 'https://phqhppdznttqskolfscp.supabase.co'
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBocWhwcGR6bnR0cXNrb2xmc2NwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk0NTYzOTksImV4cCI6MjA5NTAzMjM5OX0.N3Xvw8NPriaPf-msJokVsTKfYUNj5zYIdQmxt3k84uA'

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
    detectSessionInUrl: true,
  },
})

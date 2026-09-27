create table if not exists public.tableia_photo_sessions (
    session_id text primary key,
    image_base64 text,
    image_mime_type text,
    created_at timestamptz not null default now(),
    expires_at timestamptz not null default (now() + interval '10 minutes'),
    constraint tableia_photo_session_id_format
        check (session_id ~ '^[a-f0-9]{64}$'),
    constraint tableia_photo_mime_type
        check (
            image_mime_type is null
            or image_mime_type in (
                'image/jpeg',
                'image/png',
                'image/webp'
            )
        )
);

create index if not exists tableia_photo_sessions_expires_at_idx
    on public.tableia_photo_sessions (expires_at);

alter table public.tableia_photo_sessions
    enable row level security;

revoke all
    on table public.tableia_photo_sessions
    from anon, authenticated;

comment on table public.tableia_photo_sessions is
    'Relais temporaire et à lecture unique pour les photos TableIA.';

-- ============================================================
-- IFS Japan Visit Scheduler — Supabase Schema
-- Supabase Dashboard > SQL Editor にコピー＆実行してください
-- ============================================================

-- 来訪者テーブル（週ごとに管理）
create table if not exists visitors (
  id          uuid primary key default gen_random_uuid(),
  week_start  date not null,
  name        text not null,
  name_en     text,
  role        text,
  color_idx   integer default 0,
  created_at  timestamptz default now()
);

-- ミーティングテーブル
create table if not exists meetings (
  id            uuid primary key default gen_random_uuid(),
  week_start    date not null,
  day_index     integer not null check (day_index between 0 and 4), -- 0=月 〜 4=金
  start_time    numeric not null,  -- 9.5 = 9:30
  end_time      numeric not null,
  title         text not null,
  title_en      text,
  status        text not null default 'tentative'
                  check (status in ('confirmed', 'tentative', 'travel')),
  owner         text not null default 'TY',
  visitor_id    uuid references visitors(id) on delete set null,
  visitor_scope text,              -- 'all' の場合は全来訪者
  attendees     text[] default '{}',
  notes         text,
  briefing      text,
  created_at    timestamptz default now(),
  updated_at    timestamptz default now()
);

-- 週ベースのクエリ用インデックス
create index if not exists meetings_week_idx  on meetings(week_start);
create index if not exists visitors_week_idx  on visitors(week_start);

-- updated_at 自動更新トリガー
create or replace function update_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

drop trigger if exists meetings_updated_at on meetings;
create trigger meetings_updated_at
  before update on meetings
  for each row execute function update_updated_at();

-- Realtime 有効化（supabase_realtime publication に追加）
alter publication supabase_realtime add table meetings;
alter publication supabase_realtime add table visitors;

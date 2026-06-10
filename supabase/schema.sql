-- PT-Lovable — Supabase Schema
-- Proyecto: https://kanzirszawsgwpltdsyr.supabase.co

CREATE TABLE workplaces (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name text NOT NULL,
  client_name text,
  municipality text,
  created_at timestamptz DEFAULT now()
);

CREATE TABLE fw_applications (
  application_id uuid PRIMARY KEY,
  created_on timestamptz,
  current_stage text,
  last_modified timestamptz,
  prospect_id text,
  workplace_id text,
  workplace_name text,
  external_vacancy_id text,
  worker_id text,
  email text,
  full_name text,
  phone_number text,
  application_url text,
  cv_uploaded boolean DEFAULT false,
  cv_url text,
  imported_at timestamptz DEFAULT now()
);

CREATE TABLE status_vacancies (
  id uuid PRIMARY KEY,
  client_name text,
  workplace_name text,
  workplace_id text,
  position_name text,
  job_starts_at date,
  workers_requested integer DEFAULT 0,
  status text,
  created_at timestamptz,
  shift_pattern text,
  flow_version numeric,
  applicants_count integer DEFAULT 0,
  vr_url text,
  imported_at timestamptz DEFAULT now()
);

CREATE TABLE cache_shortlisted (
  lead_uid uuid PRIMARY KEY,
  month date,
  week date,
  day date,
  sourcing_mode text,
  vacancy_request_uid text,
  client_name text,
  position_name text,
  municipality text,
  client_agency text,
  lead_status text,
  lang_code text DEFAULT 'pt_PT',
  lead_created_at timestamptz,
  updated_at timestamptz,
  external_source_type text,
  worker_profile_id text,
  ats_id text,
  candidate_id text,
  prospect_uid text,
  first_name text,
  last_name text,
  phone text,
  email text,
  applied_on timestamptz,
  called_on timestamptz,
  contacted_on timestamptz,
  is_hired_same_client boolean DEFAULT false,
  is_hired_different_client boolean DEFAULT false,
  has_firstwork boolean DEFAULT false,
  clara_score numeric,
  imported_at timestamptz DEFAULT now()
);

CREATE TABLE placements (
  placement_id uuid PRIMARY KEY,
  country_code text DEFAULT 'PT',
  client_id text,
  workplace_id text,
  worker_id text,
  starts_at date,
  status text,
  email text,
  phone text,
  worker_shift_planned integer DEFAULT 0,
  worker_shift_and_clocked integer DEFAULT 0,
  imported_at timestamptz DEFAULT now()
);

CREATE TABLE team_mapping (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  workplace_id uuid REFERENCES workplaces(id),
  account_manager_email text NOT NULL,
  manager_type text,
  created_at timestamptz DEFAULT now()
);

CREATE TABLE candidate_comments (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  lead_uid uuid,
  author_email text,
  body text NOT NULL,
  created_at timestamptz DEFAULT now()
);

-- Desactivar RLS para desarrollo
ALTER TABLE status_vacancies DISABLE ROW LEVEL SECURITY;
ALTER TABLE fw_applications DISABLE ROW LEVEL SECURITY;
ALTER TABLE cache_shortlisted DISABLE ROW LEVEL SECURITY;
ALTER TABLE placements DISABLE ROW LEVEL SECURITY;
ALTER TABLE workplaces DISABLE ROW LEVEL SECURITY;
ALTER TABLE team_mapping DISABLE ROW LEVEL SECURITY;
ALTER TABLE candidate_comments DISABLE ROW LEVEL SECURITY;

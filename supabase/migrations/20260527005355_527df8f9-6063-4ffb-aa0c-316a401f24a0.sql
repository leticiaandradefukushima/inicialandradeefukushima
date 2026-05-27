
CREATE TABLE public.modelos_peticoes (
  tese TEXT PRIMARY KEY,
  conteudo TEXT NOT NULL DEFAULT '',
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.modelos_peticoes TO anon, authenticated;
GRANT ALL ON public.modelos_peticoes TO service_role;
ALTER TABLE public.modelos_peticoes ENABLE ROW LEVEL SECURITY;
CREATE POLICY "public read modelos" ON public.modelos_peticoes FOR SELECT TO anon, authenticated USING (true);
CREATE POLICY "public insert modelos" ON public.modelos_peticoes FOR INSERT TO anon, authenticated WITH CHECK (true);
CREATE POLICY "public update modelos" ON public.modelos_peticoes FOR UPDATE TO anon, authenticated USING (true) WITH CHECK (true);
CREATE POLICY "public delete modelos" ON public.modelos_peticoes FOR DELETE TO anon, authenticated USING (true);

CREATE TABLE public.jurisprudencias (
  uf TEXT NOT NULL,
  categoria TEXT NOT NULL,
  conteudo TEXT NOT NULL DEFAULT '',
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  PRIMARY KEY (uf, categoria)
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.jurisprudencias TO anon, authenticated;
GRANT ALL ON public.jurisprudencias TO service_role;
ALTER TABLE public.jurisprudencias ENABLE ROW LEVEL SECURITY;
CREATE POLICY "public read juris" ON public.jurisprudencias FOR SELECT TO anon, authenticated USING (true);
CREATE POLICY "public insert juris" ON public.jurisprudencias FOR INSERT TO anon, authenticated WITH CHECK (true);
CREATE POLICY "public update juris" ON public.jurisprudencias FOR UPDATE TO anon, authenticated USING (true) WITH CHECK (true);
CREATE POLICY "public delete juris" ON public.jurisprudencias FOR DELETE TO anon, authenticated USING (true);

CREATE OR REPLACE FUNCTION public.tg_touch_updated_at() RETURNS TRIGGER AS $$
BEGIN NEW.updated_at = now(); RETURN NEW; END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER touch_modelos BEFORE UPDATE ON public.modelos_peticoes
  FOR EACH ROW EXECUTE FUNCTION public.tg_touch_updated_at();
CREATE TRIGGER touch_juris BEFORE UPDATE ON public.jurisprudencias
  FOR EACH ROW EXECUTE FUNCTION public.tg_touch_updated_at();

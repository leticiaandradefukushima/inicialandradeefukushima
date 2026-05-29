-- Garante uma única linha por tese em modelos_peticoes
DELETE FROM public.modelos_peticoes a
USING public.modelos_peticoes b
WHERE a.ctid < b.ctid AND a.tese = b.tese;

ALTER TABLE public.modelos_peticoes
  DROP CONSTRAINT IF EXISTS modelos_peticoes_tese_key;
ALTER TABLE public.modelos_peticoes
  ADD CONSTRAINT modelos_peticoes_tese_key UNIQUE (tese);

-- Garante uma única linha por (uf, categoria) em jurisprudencias
DELETE FROM public.jurisprudencias a
USING public.jurisprudencias b
WHERE a.ctid < b.ctid AND a.uf = b.uf AND a.categoria = b.categoria;

ALTER TABLE public.jurisprudencias
  DROP CONSTRAINT IF EXISTS jurisprudencias_uf_categoria_key;
ALTER TABLE public.jurisprudencias
  ADD CONSTRAINT jurisprudencias_uf_categoria_key UNIQUE (uf, categoria);
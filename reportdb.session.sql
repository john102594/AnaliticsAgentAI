WITH UltimoDia AS (
  -- 1. Identificamos cuál fue el último día registrado de cada mes
  SELECT strftime('%Y-%m', fecha) AS mes_ano,
    MAX(fecha) AS fecha_corte
  FROM Historico
  WHERE proceso = 'IMPRESION'
  GROUP BY mes_ano
) -- 2. Sumamos todos los registros que coincidan con esas fechas de cierre
SELECT u.mes_ano,
  u.fecha_corte,
  CAST(SUM(h.prod_metros) AS INT) AS TotalProdMts_Cierre,
  CAST(SUM(h.mts_std) AS INT) AS TotalMtsStd_Cierre,
  ROUND(
    (
      SUM(h.prod_metros) * 1.0 / NULLIF(SUM(h.mts_std), 0)
    ) * 100,
    2
  ) AS EficienciaCierre
FROM UltimoDia u
  JOIN Historico h ON h.fecha = u.fecha_corte
WHERE h.proceso = 'IMPRESION'
GROUP BY u.mes_ano,
  u.fecha_corte
ORDER BY u.mes_ano ASC;
-- Ver los días de diciembre 2024 que existen en Historico para IMPRESION
SELECT fecha,
  turno,
  COUNT(*) as maquinas
FROM Historico
WHERE proceso = 'IMPRESION'
  AND fecha LIKE '2024-12%'
GROUP BY fecha,
  turno
ORDER BY fecha,
  turno;
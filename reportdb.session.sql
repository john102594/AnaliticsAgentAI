WITH Resumen AS (
  SELECT strftime('%Y-%m', fecha) AS mes_ano,
    SUM(prod_metros_turno) AS ProdMts,
    SUM(mts_std_turno) AS MtsStd
  FROM Historico
  WHERE proceso = 'IMPRESION'
  GROUP BY mes_ano
)
SELECT mes_ano,
  -- Convertimos a entero para quitar los decimales .0
  CAST(ProdMts AS INT) AS ProdMts,
  CAST(MtsStd AS INT) AS MtsStd,
  -- Calculamos con 2 decimales (ej. 0.85 o 85.42)
  ROUND((ProdMts * 1.0 / NULLIF(MtsStd, 0)) * 100, 2) AS PorcentajeEficiencia
FROM Resumen
ORDER BY mes_ano ASC;
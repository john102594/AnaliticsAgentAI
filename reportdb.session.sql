SELECT proceso,
  maquina,
  prod_metros_turno,
  fecha,
  turno
FROM Historico
WHERE Fecha = "2026-03-11"
  AND turno = 2
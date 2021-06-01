CREATE VIEW EXO_VENTAS_ANUAL
 AS 
 SELECT "SlpCode", "ANNO", SUM("Importe") "Importe"
 FROM (
SELECT O."SlpCode",year(O."TaxDate") "ANNO", (I."INMPrice"*I."Quantity") "Importe"
FROM OINV O 
INNER JOIN INV1 I ON O."DocEntry"=I."DocEntry"
WHERE "DocStatus"='C'
) t
GROUP BY "SlpCode", "ANNO"
ORDER BY "SlpCode", "ANNO"
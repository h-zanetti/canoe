SELECT 
    abbreviation,
--    networks
    total_imp,
--    rate_card,
    invoice_num
FROM HAGUIAR.invoice_generator_info
ORDER BY 1
;

SELECT max(invoice_num) FROM HAGUIAR.invoice_generator_info;

SET DEFINE OFF
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 566131963, invoice_num = 8515, rate_card = 0.99 WHERE abbreviation = 'A&E';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 2266094586, invoice_num = 8516, rate_card = 0.61 WHERE abbreviation = 'ABC';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 414075753, invoice_num = 8517, rate_card = 0.99 WHERE abbreviation = 'AMC';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 670676843, invoice_num = 8518, rate_card = 0.85 WHERE abbreviation = 'CBS';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 98374754, invoice_num = 8519, rate_card = 1.28 WHERE abbreviation = 'CW';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 891521, invoice_num = 8520, rate_card = 1.42 WHERE abbreviation = 'CROWN';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 1865709964, invoice_num = 8521, rate_card = 0.71 WHERE abbreviation = 'DISCOVERY';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 4530207, invoice_num = 8522, rate_card = 1.05 WHERE abbreviation = 'EPIX';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 1031152348, invoice_num = 8523, rate_card = 0.71 WHERE abbreviation = 'FOX';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 0, invoice_num = 8524, rate_card = 1.05 WHERE abbreviation = 'KIDGENIUS';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 2506626, invoice_num = 8525, rate_card = 1.05 WHERE abbreviation = 'KABILLION';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 36293873, invoice_num = 8527, rate_card = 1.28 WHERE abbreviation = 'MC';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 2956283115, invoice_num = 8528, rate_card = 0.61 WHERE abbreviation = 'NBC';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 65759, invoice_num = 8529, rate_card = 1.05 WHERE abbreviation = 'REELZ';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 608766, invoice_num = 8530, rate_card = 1.42 WHERE abbreviation = 'SONY';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 65194281, invoice_num = 8531, rate_card = 1.05 WHERE abbreviation = 'STARZ';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 2200256, invoice_num = 8532, rate_card = 1.05 WHERE abbreviation = 'TVONE';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 1875795105, invoice_num = 8533, rate_card = 0.71 WHERE abbreviation = 'TURNER';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 3156993, invoice_num = 8534, rate_card = 1.42 WHERE abbreviation = 'UNIVISION';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 1526479060, invoice_num = 8535, rate_card = 0.71 WHERE abbreviation = 'VIACOM';
commit;

SELECT os.PROGRAMMER, sum(os.IMPRESSIONS) FROM OPERATIONS.OPS_STAT_ALL os
WHERE os.EVENT_DATE >= '01-JAN-19'
AND os.EVENT_DATE <= '30-JUN-19'
GROUP BY os.PROGRAMMER ORDER BY 1
;
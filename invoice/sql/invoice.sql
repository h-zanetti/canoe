SELECT 
    abbreviation,
    networks
--    total_imp,
--    rate_card,
--    invoice_num
FROM HAGUIAR.invoice_generator_info
ORDER BY 1
;

SELECT max(invoice_num) FROM HAGUIAR.invoice_generator_info;

SET DEFINE OFF
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 453071753, invoice_num = 8492, rate_card = 0.99 WHERE abbreviation = 'A&E';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 1823113536, invoice_num = 8493, rate_card = 0.71 WHERE abbreviation = 'ABC';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 343267253, invoice_num = 8494, rate_card = 1.13 WHERE abbreviation = 'AMC';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 590720623, invoice_num = 8495, rate_card = 0.99 WHERE abbreviation = 'CBS';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 722996, invoice_num = 8497, rate_card = 1.42 WHERE abbreviation = 'CROWN';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 86407492, invoice_num = 8496, rate_card = 1.28 WHERE abbreviation = 'CW';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 1261807483, invoice_num = 8498, rate_card = 0.71 WHERE abbreviation = 'DISCOVERY';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 4484748, invoice_num = 8499, rate_card = 1.05 WHERE abbreviation = 'EPIX';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 953197184, invoice_num = 8500, rate_card = 0.71 WHERE abbreviation = 'FOX';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 1864940, invoice_num = 8502, rate_card = 1.05 WHERE abbreviation = 'KABILLION';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 0, invoice_num = 8501, rate_card = 1.05 WHERE abbreviation = 'KIDGENIUS';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 29981943, invoice_num = 8504, rate_card = 1.28 WHERE abbreviation = 'MC';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 2415073971, invoice_num = 8505, rate_card = 0.61 WHERE abbreviation = 'NBC';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 10726, invoice_num = 8506, rate_card = 1.05 WHERE abbreviation = 'REELZ';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 531796, invoice_num = 8507, rate_card = 1.42 WHERE abbreviation = 'SONY';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 52193478, invoice_num = 8508, rate_card = 1.05 WHERE abbreviation = 'STARZ';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 3115330, invoice_num = 8513, rate_card = 1.42 WHERE abbreviation = 'UNIVISION';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 1508326239, invoice_num = 8515, rate_card = 0.71 WHERE abbreviation = 'TURNER';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 1215047, invoice_num = 8509, rate_card = 1.05 WHERE abbreviation = 'TVONE';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 1211605190, invoice_num = 8514, rate_card = 0.71 WHERE abbreviation = 'VIACOM';

commit;

SELECT os.PROGRAMMER, sum(os.IMPRESSIONS) FROM OPERATIONS.OPS_STAT_ALL os
WHERE os.EVENT_DATE >= '01-JAN-19'
AND os.EVENT_DATE <= '30-JUN-19'
GROUP BY os.PROGRAMMER ORDER BY 1
;
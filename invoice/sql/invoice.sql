SELECT 
    abbreviation,
--    networks
    total_imp,
    rate_card,
    invoice_num
FROM HAGUIAR.invoice_generator_info
ORDER BY 1
;

SELECT max(invoice_num) FROM HAGUIAR.invoice_generator_info;

SET DEFINE OFF
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  676370843, invoice_num = 8515, rate_card = 0.85 WHERE abbreviation = 'A&E';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  2640507134, invoice_num = 8516, rate_card = 0.61 WHERE abbreviation = 'ABC';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  489735307, invoice_num = 8517, rate_card = 0.99 WHERE abbreviation = 'AMC';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  726105761, invoice_num = 8518, rate_card = 0.85 WHERE abbreviation = 'CBS';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  104159812, invoice_num = 8519, rate_card = 1.28 WHERE abbreviation = 'CW';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  1004954, invoice_num = 8520, rate_card = 1.42 WHERE abbreviation = 'CROWN';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  1865709964, invoice_num = 8521, rate_card = 0.71 WHERE abbreviation = 'DISCOVERY';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  4565312, invoice_num = 8522, rate_card = 1.05 WHERE abbreviation = 'EPIX';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  1098629380, invoice_num = 8523, rate_card = 0.71 WHERE abbreviation = 'FOX';
UPDATE HAGUIAR.invoice_generator_info SET total_imp = 0, invoice_num = 8524, rate_card = 1.05 WHERE abbreviation = 'KIDGENIUS';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  3164386, invoice_num = 8525, rate_card = 1.05 WHERE abbreviation = 'KABILLION';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  42822526, invoice_num = 8527, rate_card = 1.28 WHERE abbreviation = 'MC';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  3414007796, invoice_num = 8528, rate_card = 0.58 WHERE abbreviation = 'NBC';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  143297, invoice_num = 8529, rate_card = 1.05 WHERE abbreviation = 'REELZ';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  804311, invoice_num = 8530, rate_card = 1.42 WHERE abbreviation = 'SONY';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  80075521, invoice_num = 8531, rate_card = 1.05 WHERE abbreviation = 'STARZ';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  4881498, invoice_num = 8532, rate_card = 1.05 WHERE abbreviation = 'TVONE';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  2255002579, invoice_num = 8533, rate_card = 0.61 WHERE abbreviation = 'TURNER';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  4083567, invoice_num = 8534, rate_card = 1.42 WHERE abbreviation = 'UNIVISION';
UPDATE HAGUIAR.invoice_generator_info SET total_imp =  1949222409, invoice_num = 8557, rate_card = 0.71 WHERE abbreviation = 'VIACOM';
commit;


UPDATE HAGUIAR.invoice_generator_info 
SET networks = 'Bravo, E!, NBC Broadcast, Oxygen, Universal Kids, Syfy, Telemundo, USA, NBC Sports, NBC News, NBC Universo, MSNBC, CNBC, Golf Channel' 
WHERE abbreviation = 'NBC'
;
commit;


SELECT os.PROGRAMMER, sum(os.IMPRESSIONS) FROM OPERATIONS.OPS_STAT_ALL os
WHERE os.EVENT_DATE >= '01-JAN-19'
AND os.EVENT_DATE <= '30-JUN-19'
GROUP BY os.PROGRAMMER ORDER BY 1
;
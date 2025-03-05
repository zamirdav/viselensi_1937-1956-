SELECT atoc_o2_1.r07,atoc_o4_1.r07 FROM zoma.atoc_o4_1 
LEFT JOIN zoma.atoc_o2_1
ON atoc_o4_1.AT_0 = atoc_o2_1.AT_0 
where atoc_o2_1.r07 <> atoc_o4_1.r07 ;
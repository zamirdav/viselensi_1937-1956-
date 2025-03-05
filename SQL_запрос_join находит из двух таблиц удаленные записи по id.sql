SELECT * FROM zoma.toc_o2_1 
LEFT JOIN zoma.atoc_o2_1
ON toc_o2_1.AT_0 = atoc_o2_1.AT_0 
where atoc_o2_1.AT_0 is null ;
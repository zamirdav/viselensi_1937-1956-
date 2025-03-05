SELECT count(zoma.atoc_n7_1.AT_0) FROM zoma.atoc_n7_1
LEFT JOIN zoma.atoc_n1_1
ON atoc_n7_1.AT_0 = atoc_n1_1.AT_0 
where atoc_n1_1.AT_0 is null or atoc_n7_1.AT_0 is null or (
zoma.atoc_n1_1.AT_0 = zoma.atoc_n7_1.AT_0 and (
             zoma.atoc_n1_1.rO <> zoma.atoc_n7_1.rO or 
             zoma.atoc_n1_1.rG <> zoma.atoc_n7_1.rG or
             zoma.atoc_n1_1.rN <> zoma.atoc_n7_1.rN )
	);

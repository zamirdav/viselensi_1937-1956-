SELECT count(zoma.atoc_n7_20.AT_0) FROM zoma.atoc_n7_20 
LEFT JOIN zoma.atoc_n1_17
ON atoc_n7_20.AT_0 = atoc_n1_17.AT_0 
where zoma.atoc_n1_17.AT_0 = zoma.atoc_n7_20.AT_0 and (
             zoma.atoc_n1_17.rFA <> zoma.atoc_n7_20.rFA or 
             zoma.atoc_n1_17.rIM <> zoma.atoc_n7_20.rIM or
             zoma.atoc_n1_17.rOT <> zoma.atoc_n7_20.rOT or 
             zoma.atoc_n1_17.rDR <> zoma.atoc_n7_20.rDR );

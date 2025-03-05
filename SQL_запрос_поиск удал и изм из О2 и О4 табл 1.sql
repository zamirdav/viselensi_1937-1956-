SELECT count(zoma.atoc_o4_1.AT_0) FROM zoma.atoc_o4_1
LEFT JOIN zoma.atoc_o2_1
ON atoc_o4_1.AT_0 = atoc_o2_1.AT_0 
where atoc_o2_1.AT_0 is null or atoc_o2_1.AT_0 is null or (
zoma.atoc_o2_1.AT_0 = zoma.atoc_o4_1.AT_0 and (
             zoma.atoc_o2_1.r07 <> zoma.atoc_o4_1.r07 or 
             zoma.atoc_o2_1.r08 <> zoma.atoc_o4_1.r08 or
             zoma.atoc_o2_1.r09 <> zoma.atoc_o4_1.r09 or 
             zoma.atoc_o2_1.r12 <> zoma.atoc_o4_1.r12 )
);
/*SELECT a.*
FROM zoma.atoc_ai_1 a
JOIN (SELECT rO,rG,rN,rAR2,rAr12  , COUNT(*)
FROM zoma.atoc_ai_1 
GROUP BY rO,rG,rN,rAR2,rAr12
HAVING count(*) > 1 ) b
ON a.rO = b.rO
AND a.rG = b.rG
AND a.rN = b.rN
AND a.rar2 = b.rar2
ORDER BY a.rO
*/
select a.* from atoc_ai_1 a
         join  
           (select rO,rG,rN
              from atoc_ai_1 b 
              where ((rO,rG,rN) not in 
                (select (rO,rG,rN, max(AT_K_D))
                   from atoc_ai_1  
                group by (rO,rG,rN)))
           )  as c 
         using (rO,rG,rN)  
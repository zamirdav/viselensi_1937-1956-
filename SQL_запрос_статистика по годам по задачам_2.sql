DROP TABLE IF EXISTS pr2_0; /* копирует в новую таблицу pr2 */
/*CREATE TABLE pr2(bok VARCHAR(11),shap int);/*"  'создаем табл*/
CREATE TABLE pr2_0(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_0(bok,shap) 
/*SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2025' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2025'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2025'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2025'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2025'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2025'union
*/SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2025';
DROP TABLE IF EXISTS pr2; /* копирует в новую таблицу pr2 */
/*CREATE TABLE pr2(bok VARCHAR(11),shap int);/*"  'создаем табл*/
CREATE TABLE pr2(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2(bok,shap) 
/*SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2024' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2024'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2024'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2024'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2024'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2024'union
*/SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2024';

DROP TABLE IF EXISTS pr2_1; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_1(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_1(bok,shap) 
/*(SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2023' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2023' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2023' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2023' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2023' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2023' union
*/SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2023');
DROP TABLE IF EXISTS pr2_2; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_2(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_2(bok,shap) 
/*(SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2022' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2022' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2022' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2022' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2022' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2022' union
*/SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2022');
DROP TABLE IF EXISTS pr2_3; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_3(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_3(bok,shap) 
/*(SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2021' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2021' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2021' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2021' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2021' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2021' union
*/SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2021');
DROP TABLE IF EXISTS pr2_4; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_4(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_4(bok,shap) 
/*(SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2020' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2020' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2020' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2020' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2020' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2020' union
*/SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2020');

DROP TABLE IF EXISTS pr2_pr; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_pr(bok VARCHAR(11), год25 int, год24 int,год23 int,год22 int,год21 int,год20 int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_pr(bok,год25,год24,год23,год22,год21,год20)
SELECT pr2_0.bok,pr2_0.shap,pr2.bok,pr2_1.shap,pr2_2.shap,pr2_3.shap,pr2_4.shap FROM pr2
    right JOIN pr2_1 ON pr2_1.bok=pr2.bok
    right JOIN pr2_2 ON pr2_2.bok=pr2_1.bok
    right JOIN pr2_3 ON pr2_3.bok=pr2_2.bok
    right JOIN pr2_4 ON pr2_4.bok=pr2_3.bok;
    SELECT * FROM pr2_pr;
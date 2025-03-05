DROP TABLE IF EXISTS pr2_1; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_1(bok VARCHAR(11),shap int);/*"  'создаем табл*/
/*CREATE TEMPORARY TABLE pr2(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_1(bok,shap) 
SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2024' and month(Dat)='1' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2024' and month(Dat)='1'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2024' and month(Dat)='1'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2024' and month(Dat)='1'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2024' and month(Dat)='1'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2024' and month(Dat)='1'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2024' and month(Dat)='1' ;

DROP TABLE IF EXISTS pr2_2; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_2(bok VARCHAR(11),shap int);/*"  'создаем табл*/
/*CREATE TEMPORARY TABLE pr2(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_2(bok,shap) 
SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2024' and month(Dat)='2' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2024' and month(Dat)='2'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2024' and month(Dat)='2'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2024' and month(Dat)='2'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2024' and month(Dat)='2'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2024' and month(Dat)='2'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2024' and month(Dat)='2';

DROP TABLE IF EXISTS pr2_3; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_3(bok VARCHAR(11),shap int);/*"  'создаем табл*/
/*CREATE TEMPORARY TABLE pr3(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_3(bok,shap) 
SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2024' and month(Dat)='3' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2024' and month(Dat)='3'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2024' and month(Dat)='3'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2024' and month(Dat)='3'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2024' and month(Dat)='3'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2024' and month(Dat)='3'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2024' and month(Dat)='3';
DROP TABLE IF EXISTS pr2_4; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_4(bok VARCHAR(11),shap int);/*"  'создаем табл*/
/*CREATE TEMPORARY TABLE pr2(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_4(bok,shap) 
SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2024' and month(Dat)='4' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2024' and month(Dat)='4'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2024' and month(Dat)='4'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2024' and month(Dat)='4'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2024' and month(Dat)='4'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2024' and month(Dat)='4'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2024' and month(Dat)='4';
DROP TABLE IF EXISTS pr2_5; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_5(bok VARCHAR(11),shap int);/*"  'создаем табл*/
/*CREATE TEMPORARY TABLE pr2(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_5(bok,shap) 
SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2024' and month(Dat)='5' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2024' and month(Dat)='5'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2024' and month(Dat)='5'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2024' and month(Dat)='5'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2024' and month(Dat)='5'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2024' and month(Dat)='5'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2024' and month(Dat)='5';
DROP TABLE IF EXISTS pr2_6; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_6(bok VARCHAR(11),shap int);/*"  'создаем табл*/
/*CREATE TEMPORARY TABLE pr2(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_6(bok,shap) 
SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2024' and month(Dat)='6' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2024' and month(Dat)='6'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2024' and month(Dat)='6'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2024' and month(Dat)='6'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2024' and month(Dat)='6'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2024' and month(Dat)='6'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2024' and month(Dat)='6';
DROP TABLE IF EXISTS pr2_7; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_7(bok VARCHAR(11),shap int);/*"  'создаем табл*/
/*CREATE TEMPORARY TABLE pr2(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_7(bok,shap) 
SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2024' and month(Dat)='7'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2024' and month(Dat)='7'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2024' and month(Dat)='7'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2024' and month(Dat)='7'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2024' and month(Dat)='7'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2024' and month(Dat)='7'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2024' and month(Dat)='7';
DROP TABLE IF EXISTS pr2_8; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_8(bok VARCHAR(11),shap int);/*"  'создаем табл*/
/*CREATE TEMPORARY TABLE pr2(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_8(bok,shap) 
SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2024' and month(Dat)='8' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2024' and month(Dat)='8'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2024' and month(Dat)='8'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2024' and month(Dat)='8'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2024' and month(Dat)='8'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2024' and month(Dat)='8'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2024' and month(Dat)='8';
DROP TABLE IF EXISTS pr2_9; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_9(bok VARCHAR(11),shap int);/*"  'создаем табл*/
/*CREATE TEMPORARY TABLE pr2(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_9(bok,shap) 
SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2024' and month(Dat)='9' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2024' and month(Dat)='9'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2024' and month(Dat)='9'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2024' and month(Dat)='9'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2024' and month(Dat)='9'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2024' and month(Dat)='9'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2024' and month(Dat)='9';
DROP TABLE IF EXISTS pr2_10; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_10(bok VARCHAR(11),shap int);/*"  'создаем табл*/
/*CREATE TEMPORARY TABLE pr2(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_10(bok,shap) 
SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2024' and month(Dat)='10' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2024' and month(Dat)='10'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2024' and month(Dat)='10'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2024' and month(Dat)='10'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2024' and month(Dat)='10'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2024' and month(Dat)='10'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2024' and month(Dat)='10';
DROP TABLE IF EXISTS pr2_11; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_11(bok VARCHAR(11),shap int);/*"  'создаем табл*/
/*CREATE TEMPORARY TABLE pr2(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_11(bok,shap) 
SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2024' and month(Dat)='11' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2024' and month(Dat)='11'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2024' and month(Dat)='11'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2024' and month(Dat)='11'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2024' and month(Dat)='11'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2024' and month(Dat)='11'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2024' and month(Dat)='11';
DROP TABLE IF EXISTS pr2_12; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_12(bok VARCHAR(11),shap int);/*"  'создаем табл*/
/*CREATE TEMPORARY TABLE pr2(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_12(bok,shap) 
SELECT zad,count(id) FROM zoma.stat_ip where zad="o2" and year(Dat)='2024' and month(Dat)='12' union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n1" and year(Dat)='2024' and month(Dat)='12'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="n2" and year(Dat)='2024' and month(Dat)='12'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="pr" and year(Dat)='2024' and month(Dat)='12'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="rz" and year(Dat)='2024' and month(Dat)='12'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="ai" and year(Dat)='2024' and month(Dat)='12'union
SELECT zad,count(id) FROM zoma.stat_ip where zad="sp" and year(Dat)='2024' and month(Dat)='12';

DROP TABLE IF EXISTS pr2_vs; /* копирует в новую таблицу pr2 */
CREATE TABLE pr2_vs(bok VARCHAR(11),m1 int,m2 int,m3 int,m4 int,m5 int,m6 int,m7 int,m8 int,m9 int,m10 int,m11 int,m12 int);/*"  'создаем табл*/
/*CREATE TEMPORARY TABLE pr2(bok VARCHAR(11),shap int) ENGINE=MEMORY;/*"  'создаем табл*/
insert into pr2_vs(bok,m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,m11,m12) 
SELECT pr2_1.bok,pr2_1.shap,pr2_2.shap,pr2_3.shap,pr2_4.shap,pr2_5.shap,pr2_6.shap,pr2_7.shap,pr2_8.shap,pr2_9.shap,pr2_10.shap,pr2_11.shap,pr2_12.shap FROM pr2_1   
  left JOIN pr2_2 ON pr2_2.bok=pr2_1.bok     
  left JOIN pr2_3 ON pr2_3.bok=pr2_2.bok     
  left JOIN pr2_4 ON pr2_4.bok=pr2_3.bok     
  left JOIN pr2_5 ON pr2_5.bok=pr2_4.bok     
  left JOIN pr2_6 ON pr2_6.bok=pr2_5.bok     
  left JOIN pr2_7 ON pr2_7.bok=pr2_6.bok     
  left JOIN pr2_8 ON pr2_8.bok=pr2_7.bok
  left JOIN pr2_9 ON pr2_9.bok=pr2_8.bok     
  left JOIN pr2_10 ON pr2_10.bok=pr2_9.bok     
  left JOIN pr2_11 ON pr2_11.bok=pr2_10.bok     
  left JOIN pr2_12 ON pr2_12.bok=pr2_11.bok GROUP BY pr2_1.bok;
    /*SELECT * FROM pr2;

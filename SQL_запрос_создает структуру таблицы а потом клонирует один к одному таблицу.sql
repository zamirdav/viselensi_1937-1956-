/* создает структуру таблицы а потом клонирует один к одному таблицу/*
/*CREATE TABLE new_table LIKE original_table;
INSERT INTO new_table SELECT * FROM original_table;
*/
CREATE TABLE toc_ai_1 LIKE atoc_ai_1;
INSERT INTO toc_ai_1 SELECT * FROM atoc_ai_1;
--
-- SCHEMA: test
--
CREATE SCHEMA `test` DEFAULT CHARACTER SET utf8;
--
-- USER test
--
CREATE USER 'test'@'localhost' IDENTIFIED BY 'test';GRANT ALL PRIVILEGES ON `test` . * TO 'test'@'localhost';

--
--
--
CREATE DATABASE /*!32312 IF NOT EXISTS*/ `testlatin` /*!40100 DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci */ /*!80016 DEFAULT ENCRYPTION='N' */;
USE `testlatin`;
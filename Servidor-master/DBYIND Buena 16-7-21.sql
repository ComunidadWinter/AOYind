-- phpMyAdmin SQL Dump
-- version 4.1.12
-- http://www.phpmyadmin.net
--
-- Servidor: 127.0.0.1
-- Tiempo de generación: 16-07-2021 a las 20:19:11
-- Versión del servidor: 5.6.16
-- Versión de PHP: 5.5.11

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

--
-- Base de datos: `walter`
--

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `clanes`
--

CREATE TABLE IF NOT EXISTS `clanes` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Founder` varchar(20) DEFAULT NULL,
  `GuildName` varchar(20) DEFAULT NULL,
  `Fecha` date DEFAULT NULL,
  `Antifaccion` bigint(20) DEFAULT '0',
  `Alineacion` tinyint(4) DEFAULT '0',
  `Desc` varchar(500) DEFAULT NULL,
  `GuildNews` varchar(500) DEFAULT NULL,
  `URL` varchar(80) DEFAULT NULL,
  `Leader` varchar(20) DEFAULT NULL,
  `Codex1` varchar(80) DEFAULT NULL,
  `Codex2` varchar(80) DEFAULT NULL,
  `Codex3` varchar(80) DEFAULT NULL,
  `Codex4` varchar(80) DEFAULT NULL,
  `Codex5` varchar(80) DEFAULT NULL,
  `Codex6` varchar(80) DEFAULT NULL,
  `Codex7` varchar(80) DEFAULT NULL,
  `Codex8` varchar(80) DEFAULT NULL,
  `CantMiembros` int(11) DEFAULT '0',
  `EleccionesFinalizan` date DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Id` (`Id`)
) ENGINE=MyISAM  DEFAULT CHARSET=latin1 AUTO_INCREMENT=13 ;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `clanes_propuestas`
--

CREATE TABLE IF NOT EXISTS `clanes_propuestas` (
  `Id` bigint(20) NOT NULL AUTO_INCREMENT,
  `IdClan` int(11) DEFAULT NULL,
  `IdClanTo` int(11) DEFAULT NULL,
  `Detalle` varchar(400) DEFAULT NULL,
  `Tipo` int(11) DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Id` (`Id`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1 ;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `clanes_relaciones`
--

CREATE TABLE IF NOT EXISTS `clanes_relaciones` (
  `IdClan` int(11) NOT NULL,
  `IdClanTo` int(11) NOT NULL,
  `Relacion` int(11) DEFAULT '0',
  PRIMARY KEY (`IdClan`,`IdClanTo`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `clanes_solicitudes`
--

CREATE TABLE IF NOT EXISTS `clanes_solicitudes` (
  `Id` bigint(20) NOT NULL AUTO_INCREMENT,
  `IdClan` int(11) DEFAULT NULL,
  `Nombre` varchar(20) DEFAULT NULL,
  `Solicitud` varchar(400) DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Id` (`Id`)
) ENGINE=MyISAM  DEFAULT CHARSET=latin1 AUTO_INCREMENT=159 ;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `clanes_votos`
--

CREATE TABLE IF NOT EXISTS `clanes_votos` (
  `IdClan` int(11) NOT NULL,
  `Nombre` varchar(20) NOT NULL,
  `Voto` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`IdClan`,`Nombre`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `conexiones`
--

CREATE TABLE IF NOT EXISTS `conexiones` (
  `Id` bigint(20) NOT NULL AUTO_INCREMENT,
  `IdPj` bigint(20) NOT NULL,
  `IP` varchar(15) DEFAULT NULL,
  `Fecha` datetime DEFAULT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 AUTO_INCREMENT=1 ;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `cuentas`
--

CREATE TABLE IF NOT EXISTS `cuentas` (
  `Id` bigint(20) NOT NULL AUTO_INCREMENT,
  `Nombre` varchar(20) NOT NULL,
  `Password` varchar(32) DEFAULT NULL,
  `Email` varchar(100) DEFAULT NULL,
  `NombreApellido` varchar(150) DEFAULT NULL,
  `Direccion` varchar(250) DEFAULT NULL,
  `Ciudad` varchar(40) DEFAULT NULL,
  `Pais` varchar(2) DEFAULT NULL,
  `Telefono` varchar(20) DEFAULT NULL,
  `Pregunta` varchar(100) DEFAULT NULL,
  `Respuesta` varchar(50) DEFAULT NULL,
  `EmailAux` varchar(100) DEFAULT NULL,
  `Ban` tinyint(1) DEFAULT '0',
  `Nacimiento` date DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Nombre` (`Nombre`)
) ENGINE=MyISAM  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1340 ;

--
-- Volcado de datos para la tabla `cuentas`
--

INSERT INTO `cuentas` (`Id`, `Nombre`, `Password`, `Email`, `NombreApellido`, `Direccion`, `Ciudad`, `Pais`, `Telefono`, `Pregunta`, `Respuesta`, `EmailAux`, `Ban`, `Nacimiento`) VALUES
(1325, 'deidhara', '81dc9bdb52d04dc20036dbd8313ed055', 'as@as.com', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1326, 'natalia', '81dc9bdb52d04dc20036dbd8313ed055', 'nata@ass.com', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1327, 'pochola', '81dc9bdb52d04dc20036dbd8313ed055', 'pochola@pochola.com', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1328, 'asd', '81dc9bdb52d04dc20036dbd8313ed055', 'asas@a.sas', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1329, 'asasasasa', '457391c9c82bfdcbb4947278c0401e41', 'asasa@asa.asa', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1330, 'Rothson', '9c2dff566f9ec4d931df0d1a8266d014', 'asd@asd.com', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1331, 'Elias', '81dc9bdb52d04dc20036dbd8313ed055', 'elias@elias.com', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1332, 'VERONICA', '81dc9bdb52d04dc20036dbd8313ed055', 'VERO@VERO.COM', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1333, 'neosefirot', '7815696ecbf1c96e6894b779456d330e', 'a@a.com', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1334, 'sensui', '7815696ecbf1c96e6894b779456d330e', 'a@a.a', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1335, 'Sofiaa', '81dc9bdb52d04dc20036dbd8313ed055', 'as@asasss.com', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1336, 'sofia', '81dc9bdb52d04dc20036dbd8313ed055', 'sofia@sofia.com', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1337, 'pancho', '81dc9bdb52d04dc20036dbd8313ed055', 'as@asfddd.com', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1338, 'pablito', '81dc9bdb52d04dc20036dbd8313ed055', 'pablito@oablito.xom', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL),
(1339, 'dadito', '81dc9bdb52d04dc20036dbd8313ed055', 'dadi@dadi.com', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, NULL);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `fortalezas`
--

CREATE TABLE IF NOT EXISTS `fortalezas` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Nombre` varchar(50) DEFAULT NULL,
  `IdClan` int(11) DEFAULT NULL,
  `Fecha` datetime DEFAULT NULL,
  `X` int(11) DEFAULT NULL,
  `Y` int(11) DEFAULT NULL,
  `SpawnX` int(11) DEFAULT '0',
  `SpawnY` int(11) DEFAULT NULL,
  `NPCRey` int(11) DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Id` (`Id`)
) ENGINE=MyISAM  DEFAULT CHARSET=latin1 AUTO_INCREMENT=3 ;

--
-- Volcado de datos para la tabla `fortalezas`
--

INSERT INTO `fortalezas` (`Id`, `Nombre`, `IdClan`, `Fecha`, `X`, `Y`, `SpawnX`, `SpawnY`, `NPCRey`) VALUES
(1, 'Oeste', 11, '2018-12-15 14:50:15', 40, 1460, 18, 1304, 620),
(2, 'Este', 11, '2018-12-15 14:47:58', 1055, 1460, 1080, 1252, 619);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `penas`
--

CREATE TABLE IF NOT EXISTS `penas` (
  `Id` tinyint(20) NOT NULL AUTO_INCREMENT,
  `IdPj` bigint(20) DEFAULT NULL,
  `Razon` varchar(100) DEFAULT NULL,
  `Fecha` datetime NOT NULL,
  `IdGM` bigint(20) DEFAULT NULL,
  `Tiempo` int(11) DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Id` (`Id`)
) ENGINE=MyISAM  DEFAULT CHARSET=utf8 AUTO_INCREMENT=128 ;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `pjs`
--

CREATE TABLE IF NOT EXISTS `pjs` (
  `Id` bigint(20) unsigned NOT NULL AUTO_INCREMENT,
  `IdAccount` bigint(20) DEFAULT '0',
  `Nombre` varchar(20) DEFAULT NULL,
  `Genero` tinyint(4) DEFAULT NULL,
  `Raza` tinyint(4) DEFAULT NULL,
  `Hogar` tinyint(4) DEFAULT NULL,
  `Clase` tinyint(4) DEFAULT NULL,
  `Heading` tinyint(4) DEFAULT NULL,
  `Head` int(11) DEFAULT NULL,
  `Body` int(11) DEFAULT NULL,
  `Arma` int(11) DEFAULT NULL,
  `Escudo` int(11) DEFAULT NULL,
  `Casco` int(11) DEFAULT NULL,
  `AlaIndex` int(11) DEFAULT '0',
  `Uptime` bigint(20) DEFAULT NULL,
  `LastIP` varchar(15) DEFAULT NULL,
  `LastConnect` datetime DEFAULT NULL,
  `Map` tinyint(4) DEFAULT NULL,
  `X` int(11) DEFAULT NULL,
  `Y` int(11) DEFAULT NULL,
  `Logged` tinyint(1) DEFAULT NULL,
  `Muerto` tinyint(1) DEFAULT NULL,
  `Escondido` tinyint(1) DEFAULT NULL,
  `Hambre` tinyint(1) DEFAULT NULL,
  `Sed` tinyint(1) DEFAULT NULL,
  `Desnudo` tinyint(1) DEFAULT NULL,
  `Ban` tinyint(1) DEFAULT NULL,
  `Navegando` tinyint(1) DEFAULT NULL,
  `Envenenado` tinyint(1) DEFAULT NULL,
  `Paralizado` tinyint(1) DEFAULT NULL,
  `PerteneceReal` tinyint(1) DEFAULT NULL,
  `PerteneceCaos` tinyint(1) DEFAULT NULL,
  `Pena` int(11) DEFAULT NULL,
  `EjercitoReal` tinyint(1) DEFAULT NULL,
  `EjercitoCaos` tinyint(1) DEFAULT NULL,
  `CiudMatados` tinyint(1) DEFAULT NULL,
  `CrimMatados` tinyint(1) DEFAULT NULL,
  `rArCaos` tinyint(1) DEFAULT NULL,
  `rArReal` tinyint(1) DEFAULT NULL,
  `rExCaos` tinyint(1) DEFAULT NULL,
  `rExReal` tinyint(1) DEFAULT NULL,
  `recCaos` int(11) DEFAULT NULL,
  `recReal` int(11) DEFAULT NULL,
  `Reenlistadas` int(11) DEFAULT NULL,
  `NivelIngreso` int(11) DEFAULT NULL,
  `FechaIngreso` date DEFAULT NULL,
  `MatadosIngreso` int(11) DEFAULT NULL,
  `NextRecompensa` int(11) DEFAULT NULL,
  `At1` int(11) DEFAULT NULL,
  `At2` int(11) DEFAULT NULL,
  `At3` int(11) DEFAULT NULL,
  `At4` int(11) DEFAULT NULL,
  `At5` int(11) DEFAULT NULL,
  `Sk1` int(11) DEFAULT NULL,
  `Sk2` int(11) DEFAULT NULL,
  `Sk3` int(11) DEFAULT NULL,
  `Sk4` int(11) DEFAULT NULL,
  `Sk5` int(11) DEFAULT NULL,
  `Sk6` int(11) DEFAULT NULL,
  `Sk7` int(11) DEFAULT NULL,
  `Sk8` int(11) DEFAULT NULL,
  `Sk9` int(11) DEFAULT NULL,
  `Sk10` int(11) DEFAULT NULL,
  `Sk11` int(11) DEFAULT NULL,
  `Sk12` int(11) DEFAULT NULL,
  `Sk13` int(11) DEFAULT NULL,
  `Sk14` int(11) DEFAULT NULL,
  `Sk15` int(11) DEFAULT NULL,
  `Sk16` int(11) DEFAULT NULL,
  `Sk17` int(11) DEFAULT NULL,
  `Sk18` int(11) DEFAULT NULL,
  `Sk19` int(11) DEFAULT NULL,
  `Sk20` int(11) DEFAULT NULL,
  `Email` varchar(100) DEFAULT NULL,
  `Gld` bigint(20) DEFAULT NULL,
  `Banco` bigint(20) DEFAULT NULL,
  `MaxHP` int(11) DEFAULT NULL,
  `MinHP` int(11) DEFAULT NULL,
  `MaxSTA` int(11) DEFAULT NULL,
  `MinSTA` int(11) DEFAULT NULL,
  `MaxMAN` int(11) DEFAULT NULL,
  `MinMAN` int(11) DEFAULT NULL,
  `MaxHIT` int(11) DEFAULT NULL,
  `MinHIT` int(11) DEFAULT NULL,
  `MaxAGU` int(11) DEFAULT NULL,
  `MinAGU` int(11) DEFAULT NULL,
  `MaxHAM` int(11) DEFAULT NULL,
  `MinHAM` int(11) DEFAULT NULL,
  `SkillPtsLibres` int(11) DEFAULT NULL,
  `EXP` bigint(20) DEFAULT NULL,
  `ELV` int(11) DEFAULT NULL,
  `ELU` bigint(20) DEFAULT NULL,
  `UserMuertes` bigint(20) DEFAULT NULL,
  `MuertesPropias` bigint(20) DEFAULT '0',
  `NpcsMuertes` bigint(20) DEFAULT NULL,
  `WeaponEqpSlot` int(11) DEFAULT NULL,
  `ArmourEqpSlot` int(11) DEFAULT NULL,
  `CascoEqpSlot` int(11) DEFAULT NULL,
  `EscudoEqpSlot` int(11) DEFAULT NULL,
  `AlaEqpSlot` int(11) DEFAULT NULL,
  `BarcoSlot` int(11) DEFAULT NULL,
  `MunicionSlot` int(11) DEFAULT NULL,
  `AnilloSlot` int(11) DEFAULT NULL,
  `Rep_Asesino` bigint(20) DEFAULT NULL,
  `Rep_Bandido` bigint(20) DEFAULT NULL,
  `Rep_Burguesia` bigint(20) DEFAULT NULL,
  `Rep_Ladrones` bigint(20) DEFAULT NULL,
  `Rep_Nobles` bigint(20) DEFAULT NULL,
  `Rep_Plebe` bigint(20) DEFAULT NULL,
  `Rep_Promedio` bigint(20) DEFAULT NULL,
  `NroMascotas` int(11) DEFAULT NULL,
  `Masc1` int(11) DEFAULT NULL,
  `Masc2` int(11) DEFAULT NULL,
  `Masc3` int(11) DEFAULT NULL,
  `TrainningTime` bigint(20) DEFAULT '0',
  `H1` int(11) DEFAULT NULL,
  `H2` int(11) DEFAULT NULL,
  `H3` int(11) DEFAULT NULL,
  `H4` int(11) DEFAULT NULL,
  `H5` int(11) DEFAULT NULL,
  `H6` int(11) DEFAULT NULL,
  `H7` int(11) DEFAULT NULL,
  `H8` int(11) DEFAULT NULL,
  `H9` int(11) DEFAULT NULL,
  `H10` int(11) DEFAULT NULL,
  `H11` int(11) DEFAULT NULL,
  `H12` int(11) DEFAULT NULL,
  `H13` int(11) DEFAULT NULL,
  `H14` int(11) DEFAULT NULL,
  `H15` int(11) DEFAULT NULL,
  `H16` int(11) DEFAULT NULL,
  `H17` int(11) DEFAULT NULL,
  `H18` int(11) DEFAULT NULL,
  `H19` int(11) DEFAULT NULL,
  `H20` int(11) DEFAULT NULL,
  `H21` int(11) DEFAULT NULL,
  `H22` int(11) DEFAULT NULL,
  `H23` int(11) DEFAULT NULL,
  `H24` int(11) DEFAULT NULL,
  `H25` int(11) DEFAULT NULL,
  `H26` int(11) DEFAULT NULL,
  `H27` int(11) DEFAULT NULL,
  `H28` int(11) DEFAULT NULL,
  `H29` int(11) DEFAULT NULL,
  `H30` int(11) DEFAULT NULL,
  `H31` int(11) DEFAULT NULL,
  `H32` int(11) DEFAULT NULL,
  `H33` int(11) DEFAULT NULL,
  `H34` int(11) DEFAULT NULL,
  `H35` int(11) DEFAULT NULL,
  `InvObj1` int(11) DEFAULT NULL,
  `InvCant1` int(11) DEFAULT NULL,
  `InvEqp1` tinyint(1) DEFAULT NULL,
  `InvObj2` int(11) DEFAULT NULL,
  `InvCant2` int(11) DEFAULT NULL,
  `InvEqp2` tinyint(1) DEFAULT NULL,
  `InvObj3` int(11) DEFAULT NULL,
  `InvCant3` int(11) DEFAULT NULL,
  `InvEqp3` tinyint(1) DEFAULT NULL,
  `InvObj4` int(11) DEFAULT NULL,
  `InvCant4` int(11) DEFAULT NULL,
  `InvEqp4` tinyint(1) DEFAULT NULL,
  `InvObj5` int(11) DEFAULT NULL,
  `InvCant5` int(11) DEFAULT NULL,
  `InvEqp5` tinyint(1) DEFAULT NULL,
  `InvObj6` int(11) DEFAULT NULL,
  `InvCant6` int(11) DEFAULT NULL,
  `InvEqp6` tinyint(1) DEFAULT NULL,
  `InvObj7` int(11) DEFAULT NULL,
  `InvCant7` int(11) DEFAULT NULL,
  `InvEqp7` tinyint(1) DEFAULT NULL,
  `InvObj8` int(11) DEFAULT NULL,
  `InvCant8` int(11) DEFAULT NULL,
  `InvEqp8` tinyint(1) DEFAULT NULL,
  `InvObj9` int(11) DEFAULT NULL,
  `InvCant9` int(11) DEFAULT NULL,
  `InvEqp9` tinyint(1) DEFAULT NULL,
  `InvObj10` int(11) DEFAULT NULL,
  `InvCant10` int(11) DEFAULT NULL,
  `InvEqp10` tinyint(1) DEFAULT NULL,
  `InvObj11` int(11) DEFAULT NULL,
  `InvCant11` int(11) DEFAULT NULL,
  `InvEqp11` tinyint(1) DEFAULT NULL,
  `InvObj12` int(11) DEFAULT NULL,
  `InvCant12` int(11) DEFAULT NULL,
  `InvEqp12` tinyint(1) DEFAULT NULL,
  `InvObj13` int(11) DEFAULT NULL,
  `InvCant13` int(11) DEFAULT NULL,
  `InvEqp13` tinyint(1) DEFAULT NULL,
  `InvObj14` int(11) DEFAULT NULL,
  `InvCant14` int(11) DEFAULT NULL,
  `InvEqp14` tinyint(1) DEFAULT NULL,
  `InvObj15` int(11) DEFAULT NULL,
  `InvCant15` int(11) DEFAULT NULL,
  `InvEqp15` tinyint(1) DEFAULT NULL,
  `InvObj16` int(11) DEFAULT NULL,
  `InvCant16` int(11) DEFAULT NULL,
  `InvEqp16` tinyint(1) DEFAULT NULL,
  `InvObj17` int(11) DEFAULT NULL,
  `InvCant17` int(11) DEFAULT NULL,
  `InvEqp17` tinyint(1) DEFAULT NULL,
  `InvObj18` int(11) DEFAULT NULL,
  `InvCant18` int(11) DEFAULT NULL,
  `InvEqp18` tinyint(1) DEFAULT NULL,
  `InvObj19` int(11) DEFAULT NULL,
  `InvCant19` int(11) DEFAULT NULL,
  `InvEqp19` tinyint(1) DEFAULT NULL,
  `InvObj20` int(11) DEFAULT NULL,
  `InvCant20` int(11) DEFAULT NULL,
  `InvEqp20` tinyint(1) DEFAULT NULL,
  `InvCantidadItems` int(11) DEFAULT NULL,
  `GuildIndex` int(11) DEFAULT '0',
  `Descripcion` varchar(100) DEFAULT NULL,
  `Creado` datetime DEFAULT NULL,
  `BannedBy` varchar(20) DEFAULT NULL,
  `Voto` int(11) DEFAULT '0',
  `AspiranteA` int(11) DEFAULT '0',
  `MotivoRechazo` varchar(100) DEFAULT '',
  `Pedidos` varchar(400) DEFAULT '0',
  `Miembro` varchar(400) DEFAULT NULL,
  `Extra` varchar(100) DEFAULT NULL,
  `Penas` int(11) DEFAULT '0',
  `BanTime` date DEFAULT '2000-01-01',
  `Q1` varchar(20) DEFAULT '0',
  `Q2` varchar(20) DEFAULT '0',
  `Q3` varchar(20) DEFAULT '0',
  `Q4` varchar(20) DEFAULT '0',
  `Q5` varchar(20) DEFAULT '0',
  `Q6` varchar(20) DEFAULT '0',
  `Q7` varchar(20) DEFAULT '0',
  `Q8` varchar(20) DEFAULT '0',
  `Q9` varchar(20) DEFAULT '0',
  `Q10` varchar(20) DEFAULT '0',
  `Q11` varchar(20) DEFAULT '0',
  `Q12` varchar(20) DEFAULT '0',
  `Q13` varchar(20) DEFAULT '0',
  `Q14` varchar(20) DEFAULT '0',
  `Q15` varchar(20) DEFAULT '0',
  `QuestsDone` varchar(20) DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Nombre` (`Nombre`)
) ENGINE=MyISAM  DEFAULT CHARSET=utf8 AUTO_INCREMENT=2381 ;

--
-- Volcado de datos para la tabla `pjs`
--

INSERT INTO `pjs` (`Id`, `IdAccount`, `Nombre`, `Genero`, `Raza`, `Hogar`, `Clase`, `Heading`, `Head`, `Body`, `Arma`, `Escudo`, `Casco`, `AlaIndex`, `Uptime`, `LastIP`, `LastConnect`, `Map`, `X`, `Y`, `Logged`, `Muerto`, `Escondido`, `Hambre`, `Sed`, `Desnudo`, `Ban`, `Navegando`, `Envenenado`, `Paralizado`, `PerteneceReal`, `PerteneceCaos`, `Pena`, `EjercitoReal`, `EjercitoCaos`, `CiudMatados`, `CrimMatados`, `rArCaos`, `rArReal`, `rExCaos`, `rExReal`, `recCaos`, `recReal`, `Reenlistadas`, `NivelIngreso`, `FechaIngreso`, `MatadosIngreso`, `NextRecompensa`, `At1`, `At2`, `At3`, `At4`, `At5`, `Sk1`, `Sk2`, `Sk3`, `Sk4`, `Sk5`, `Sk6`, `Sk7`, `Sk8`, `Sk9`, `Sk10`, `Sk11`, `Sk12`, `Sk13`, `Sk14`, `Sk15`, `Sk16`, `Sk17`, `Sk18`, `Sk19`, `Sk20`, `Email`, `Gld`, `Banco`, `MaxHP`, `MinHP`, `MaxSTA`, `MinSTA`, `MaxMAN`, `MinMAN`, `MaxHIT`, `MinHIT`, `MaxAGU`, `MinAGU`, `MaxHAM`, `MinHAM`, `SkillPtsLibres`, `EXP`, `ELV`, `ELU`, `UserMuertes`, `MuertesPropias`, `NpcsMuertes`, `WeaponEqpSlot`, `ArmourEqpSlot`, `CascoEqpSlot`, `EscudoEqpSlot`, `AlaEqpSlot`, `BarcoSlot`, `MunicionSlot`, `AnilloSlot`, `Rep_Asesino`, `Rep_Bandido`, `Rep_Burguesia`, `Rep_Ladrones`, `Rep_Nobles`, `Rep_Plebe`, `Rep_Promedio`, `NroMascotas`, `Masc1`, `Masc2`, `Masc3`, `TrainningTime`, `H1`, `H2`, `H3`, `H4`, `H5`, `H6`, `H7`, `H8`, `H9`, `H10`, `H11`, `H12`, `H13`, `H14`, `H15`, `H16`, `H17`, `H18`, `H19`, `H20`, `H21`, `H22`, `H23`, `H24`, `H25`, `H26`, `H27`, `H28`, `H29`, `H30`, `H31`, `H32`, `H33`, `H34`, `H35`, `InvObj1`, `InvCant1`, `InvEqp1`, `InvObj2`, `InvCant2`, `InvEqp2`, `InvObj3`, `InvCant3`, `InvEqp3`, `InvObj4`, `InvCant4`, `InvEqp4`, `InvObj5`, `InvCant5`, `InvEqp5`, `InvObj6`, `InvCant6`, `InvEqp6`, `InvObj7`, `InvCant7`, `InvEqp7`, `InvObj8`, `InvCant8`, `InvEqp8`, `InvObj9`, `InvCant9`, `InvEqp9`, `InvObj10`, `InvCant10`, `InvEqp10`, `InvObj11`, `InvCant11`, `InvEqp11`, `InvObj12`, `InvCant12`, `InvEqp12`, `InvObj13`, `InvCant13`, `InvEqp13`, `InvObj14`, `InvCant14`, `InvEqp14`, `InvObj15`, `InvCant15`, `InvEqp15`, `InvObj16`, `InvCant16`, `InvEqp16`, `InvObj17`, `InvCant17`, `InvEqp17`, `InvObj18`, `InvCant18`, `InvEqp18`, `InvObj19`, `InvCant19`, `InvEqp19`, `InvObj20`, `InvCant20`, `InvEqp20`, `InvCantidadItems`, `GuildIndex`, `Descripcion`, `Creado`, `BannedBy`, `Voto`, `AspiranteA`, `MotivoRechazo`, `Pedidos`, `Miembro`, `Extra`, `Penas`, `BanTime`, `Q1`, `Q2`, `Q3`, `Q4`, `Q5`, `Q6`, `Q7`, `Q8`, `Q9`, `Q10`, `Q11`, `Q12`, `Q13`, `Q14`, `Q15`, `QuestsDone`) VALUES
(2356, 1326, 'natalia', 2, 1, 2, 6, 3, 70, 1, 12, 2, 2, 0, 73, '168.196.245.137', '2021-04-24 17:03:20', 1, 179, 1238, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 19, 19, 18, 18, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 19, 19, 60, 60, 50, 50, 2, 1, 100, 100, 100, 100, 10, 0, 1, 300, 0, 0, 0, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1000, 30, 172, 0, 0, 0, 0, 72, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 70, 0, 491, 100, 0, 460, 1, 1, 463, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 0, '', '2021-04-21 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2357, 1325, 'feoo', 1, 1, 1, 4, 3, 22, 21, 2, 2, 2, 0, 8249, '127.0.0.1', '2021-07-09 18:58:47', 1, 264, 872, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 19, 19, 18, 18, 20, 11, 0, 35, 5, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 23, 0, '', 0, 0, 119, 119, 235, 235, 284, 284, 41, 40, 100, 80, 100, 70, 70, 2244, 14, 23799, 0, 0, 37, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 19500, 30, 3255, 0, 0, 0, 0, 1078, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '2021-04-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '1'),
(2358, 1325, 'benjamin', 1, 2, 2, 2, 1, 108, 210, 2, 2, 2, 0, 40326, '127.0.0.1', '2021-07-11 00:37:23', 8, 65, 52, 0, 0, 0, 0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 17, 21, 20, 20, 19, 100, 0, 67, 42, 67, 0, 0, 0, 0, 14, 0, 0, 0, 0, 0, 0, 0, 0, 53, 0, '', 1284, 0, 219, 179, 450, 450, 1090, 1090, 54, 53, 100, 80, 100, 90, 82, 130561, 27, 807237, 0, 0, 174, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 88000, 30, 14672, 0, 0, 0, 0, 1043, 15, 0, 0, 0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1002, 2, 0, 160, 4, 0, 35, 1, 0, 501, 2, 0, 237, 4, 0, 402, 1, 0, 22, 9, 0, 80, 7, 0, 556, 1, 0, 1, 30, 0, 38, 487, 0, 158, 11, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 12, 0, '', '2021-04-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2359, 1327, 'pochola', 1, 1, 1, 2, 3, 24, 1, 12, 2, 2, 0, 25, '168.196.245.137', '2021-04-23 11:04:03', 1, 304, 854, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 19, 19, 18, 18, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 18, 18, 40, 40, 50, 50, 2, 1, 100, 100, 100, 100, 10, 0, 1, 300, 0, 0, 0, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1000, 30, 172, 0, 0, 0, 0, 25, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 70, 0, 491, 100, 0, 460, 1, 1, 463, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 0, '', '2021-04-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2360, 1325, 'RENATA', 2, 2, 2, 2, 4, 177, 259, 2, 2, 2, 0, 9812, '127.0.0.1', '2021-07-11 00:41:46', 1, 203, 1268, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 17, 21, 20, 20, 19, 14, 0, 38, 14, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 32, 0, '', 0, 0, 134, 134, 265, 265, 650, 610, 32, 31, 100, 20, 100, 30, 79, 3449, 16, 40221, 0, 0, 37, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 19500, 30, 3255, 0, 0, 0, 0, 887, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '2021-04-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2361, 1326, 'BENJAMINES', 1, 4, 2, 2, 3, 403, 300, 12, 2, 2, 0, 234, '168.196.245.137', '2021-04-23 13:21:32', 1, 231, 1266, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 16, 21, 22, 19, 18, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 100, 0, 19, 19, 40, 40, 50, 50, 2, 1, 100, 100, 100, 100, 10, 150, 1, 300, 0, 0, 0, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1000, 30, 172, 0, 0, 0, 0, 233, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 70, 0, 491, 100, 0, 460, 1, 1, 466, 1, 1, 127, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 7, 0, '', '2021-04-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '1-1'),
(2362, 1328, 'sensui', 1, 1, 2, 2, 1, 37, 3, 15, 2, 2, 0, 1478, '213.94.53.229', '2021-04-28 12:03:45', 1, 188, 1212, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 19, 19, 18, 18, 20, 0, 0, 0, 8, 0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 22, '', 132544, 0, 288, 288, 555, 555, 1238, 50, 68, 67, 100, 100, 100, 100, 158, 1553094, 34, 6596865, 0, 0, 5, 4, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3500, 30, 588, 0, 0, 0, 0, 1209, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 389, 1, 0, 2, 1, 0, 35, 1, 1, 556, 1, 1, 1100, 1, 0, 127, 1, 0, 43, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 7, 0, '', '2021-04-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '-1'),
(2363, 1328, 'asasa', 1, 1, 2, 2, 1, 19, 21, 2, 2, 2, 0, 496, '213.94.53.229', '2021-04-23 16:35:38', 1, 223, 1266, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 19, 19, 18, 18, 20, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 257, 0, 505, 0, 1166, 50, 64, 63, 100, 100, 100, 100, 170, 2925510, 32, 3619679, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2000, 32, 339, 0, 0, 0, 0, 382, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '2021-04-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '15-2', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2374, 1334, 'sensuito', 1, 2, 2, 2, 3, 112, 2, 2, 2, 2, 0, 521, '213.94.52.25', '2021-05-14 17:11:32', 13, 193, 360, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 17, 21, 20, 20, 19, 0, 0, 9, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3, 0, '', 0, 0, 66, 0, 130, 0, 290, 210, 14, 13, 100, 100, 100, 100, 45, 1161, 7, 2258, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2000, 30, 338, 0, 0, 0, 0, 33, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 65, 0, 491, 89, 0, 460, 1, 0, 464, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 0, '', '2021-05-14 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2364, 1329, 'sasasasa', 1, 2, 2, 2, 3, 108, 2, 12, 2, 2, 0, 251, '213.94.53.229', '2021-04-28 12:00:32', 1, 188, 1213, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 17, 21, 20, 20, 19, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 20, 20, 40, 40, 50, 50, 2, 1, 100, 100, 100, 100, 10, 0, 1, 300, 0, 0, 0, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1000, 30, 172, 0, 0, 0, 0, 251, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 70, 0, 491, 100, 0, 460, 1, 1, 464, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 0, '', '2021-04-28 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '5'),
(2365, 1330, 'Rothson', 1, 1, 2, 2, 3, 11, 1, 12, 2, 2, 0, 611, '190.244.34.31', '2021-04-28 20:42:49', 1, 188, 1235, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 19, 19, 18, 18, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 18, 18, 60, 60, 50, 50, 2, 1, 100, 90, 100, 100, 10, 0, 1, 300, 0, 0, 0, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1000, 30, 172, 0, 0, 0, 0, 612, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 70, 0, 491, 100, 0, 460, 1, 1, 463, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 0, '', '2021-04-28 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '5'),
(2368, 1325, 'pepita', 1, 5, 2, 3, 3, 302, 53, 2, 2, 2, 0, 1089, '127.0.0.1', '2021-07-07 20:35:50', 8, 40, 59, 0, 0, 0, 0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 21, 18, 16, 16, 21, 0, 0, 33, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 142, 23, 220, 220, 0, 0, 38, 37, 100, 100, 100, 100, 75, 2590, 13, 16999, 0, 0, 11, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6500, 30, 1088, 0, 0, 0, 0, 932, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '2021-05-09 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2366, 1331, 'EliasG', 1, 2, 1, 7, 4, 106, 210, 2, 2, 2, 0, 2911, '168.196.245.137', '2021-05-02 00:48:33', 1, 320, 882, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 17, 21, 20, 20, 19, 45, 0, 45, 38, 11, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, 0, 31, 0, 0, 0, '', 471656, 0, 143, 143, 315, 98, 730, 310, 36, 35, 100, 30, 100, 70, 25, 23596, 18, 67973, 0, 0, 47, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 24500, 30, 4088, 0, 0, 0, 0, 368, 2, 46, 6, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '2021-05-01 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2367, 1325, 'hola', 1, 5, 1, 3, 3, 300, 53, 2, 2, 2, 0, 1681, '127.0.0.1', '2021-07-09 18:59:41', 8, 57, 64, 0, 0, 0, 0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 21, 18, 16, 16, 21, 0, 0, 10, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 5, 0, '', 0, 0, 153, 113, 235, 235, 0, 0, 41, 40, 100, 100, 100, 100, 80, 6751, 14, 23799, 0, 0, 18, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10000, 30, 1672, 0, 0, 0, 0, 236, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '2021-05-09 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2369, 1325, 'senputo', 1, 5, 2, 3, 3, 302, 53, 2, 2, 2, 0, 963, '127.0.0.1', '2021-07-10 18:08:36', 1, 176, 1236, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 21, 18, 16, 16, 21, 0, 0, 3, 9, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 136, 136, 240, 240, 0, 0, 38, 37, 100, 100, 100, 100, 75, 750, 13, 16999, 0, 0, 17, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 9500, 30, 1588, 0, 0, 0, 0, 504, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '2021-05-09 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2370, 1325, 'pepitts', 1, 1, 2, 3, 3, 24, 21, 2, 2, 2, 0, 737, '127.0.0.1', '2021-07-09 18:53:30', 8, 79, 66, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 19, 19, 18, 18, 20, 0, 0, 28, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 23, 0, '', 197, 0, 164, 0, 265, 0, 0, 0, 47, 46, 100, 100, 100, 100, 90, 10960, 16, 40221, 0, 0, 45, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 23500, 30, 3922, 0, 0, 0, 0, 71, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '2021-05-09 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2371, 1325, 'DFGG', 1, 5, 1, 3, 1, 305, 53, 2, 2, 2, 0, 1917, '127.0.0.1', '2021-07-10 18:07:55', 8, 67, 100, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 21, 18, 16, 16, 21, 0, 0, 33, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 127, 0, 220, 0, 0, 0, 38, 37, 100, 80, 100, 80, 75, 3010, 13, 16999, 0, 2, 17, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 9500, 30, 1588, 0, 0, 0, 0, 1783, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '2021-05-09 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2372, 1332, 'VERONIQUITA', 1, 5, 3, 3, 3, 304, 53, 2, 2, 2, 0, 670, '127.0.0.1', '2021-06-23 10:35:49', 1, 300, 220, 0, 0, 0, 0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 21, 18, 16, 16, 21, 0, 0, 15, 9, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 147, 102, 220, 220, 0, 0, 38, 37, 100, 100, 100, 100, 75, 1210, 13, 16999, 0, 0, 15, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 8500, 30, 1422, 0, 0, 0, 0, 601, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '2021-05-09 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2373, 1333, 'neosefirot', 2, 3, 2, 5, 3, 278, 3, 12, 2, 2, 0, 69, '62.83.19.236', '2021-05-10 16:04:07', 1, 194, 1252, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 20, 20, 20, 15, 18, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 20, 20, 40, 40, 0, 0, 2, 1, 100, 100, 100, 100, 10, 0, 1, 300, 0, 0, 0, 4, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1000, 30, 172, 0, 0, 0, 0, 69, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 70, 0, 460, 1, 1, 465, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 5, 0, '', '2021-05-10 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2375, 1335, 'Sofiass', 2, 5, 1, 3, 3, 471, 60, 2, 2, 2, 0, 2208, '127.0.0.1', '2021-05-21 12:53:54', 1, 177, 1185, 0, 0, 0, 0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 21, 18, 16, 16, 21, 0, 0, 10, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 7, 0, '', 0, 0, 148, 135, 235, 235, 0, 0, 41, 40, 100, 80, 100, 80, 80, 1911, 14, 23799, 0, 0, 21, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 11500, 30, 1922, 0, 0, 0, 0, 1457, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '2021-05-21 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2377, 1325, 'deidhara', 1, 2, 1, 7, 3, 105, 130, 10, 2, 4, 2, 49604, '127.0.0.1', '2021-07-14 16:26:37', 1, 1000, 994, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 17, 21, 20, 20, 19, 100, 0, 0, 63, 0, 0, 0, 0, 12, 12, 0, 0, 58, 0, 0, 0, 0, 0, 62, 0, '', 996886, 0, 400, 400, 400, 400, 1010, 1010, 50, 49, 100, 100, 100, 100, 5, 234048, 5, 442928, 0, 0, 56, 18, 8, 6, 0, 16, 0, 0, 0, 0, 0, 0, 0, 29000, 694, 4949, 0, 0, 0, 0, 10706, 46, 2, 0, 0, 25, 0, 0, 0, 15, 0, 0, 0, 24, 0, 0, 0, 23, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 556, 1, 0, 662, 1, 1, 193, 5, 0, 519, 1, 1, 43, 3, 0, 3, 3, 0, 187, 1, 0, 192, 24, 0, 1, 5, 0, 511, 20, 0, 2, 6, 0, 1111, 1, 1, 80, 1, 0, 660, 1, 1, 22, 8, 0, 1110, 4, 0, 17, 0, '', '2021-06-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2376, 1325, 'renannata', 2, 2, 1, 2, 3, 174, 259, 2, 2, 2, 0, 1137, '127.0.0.1', '2021-07-12 17:07:01', 8, 71, 56, 0, 0, 0, 0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 17, 21, 20, 20, 19, 0, 0, 9, 14, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 114, 111, 220, 220, 530, 450, 26, 25, 100, 90, 100, 90, 75, 37, 13, 16999, 0, 0, 14, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 8000, 30, 1338, 0, 0, 0, 0, 88, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '2021-06-23 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2380, 1338, 'pablito', 1, 1, 1, 2, 3, 26, 1, 12, 2, 2, 0, 8, '127.0.0.1', '2021-07-10 23:58:21', 1, 304, 854, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 19, 19, 18, 18, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 19, 19, 40, 40, 50, 50, 2, 1, 100, 100, 100, 100, 10, 0, 1, 300, 0, 0, 0, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1000, 30, 172, 0, 0, 0, 0, 7, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 70, 0, 491, 100, 0, 460, 1, 1, 463, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 0, '', '2021-07-10 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2378, 1336, 'pocholas', 2, 2, 1, 2, 3, 174, 2, 12, 2, 2, 0, 113, '127.0.0.1', '2021-07-07 20:32:03', 1, 301, 870, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 17, 21, 20, 20, 19, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 19, 19, 60, 60, 50, 50, 2, 1, 100, 100, 100, 100, 10, 0, 1, 300, 0, 0, 0, 5, 6, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 1000, 30, 172, 0, 0, 0, 0, 112, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 70, 0, 491, 100, 0, 460, 1, 1, 464, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 0, '', '2021-07-07 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'),
(2379, 1337, 'amilcar', 1, 2, 1, 2, 3, 104, 2, 12, 2, 2, 0, 53, '127.0.0.1', '2021-07-07 20:34:42', 1, 307, 862, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '2000-01-01', 0, 0, 17, 21, 20, 20, 19, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 19, 19, 60, 60, 50, 50, 2, 1, 100, 100, 100, 100, 10, 0, 1, 300, 0, 0, 0, 5, 6, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 1000, 30, 172, 0, 0, 0, 0, 53, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 467, 100, 0, 468, 100, 0, 461, 70, 0, 491, 100, 0, 460, 1, 1, 464, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 0, '', '2021-07-07 00:00:00', '', 0, 0, '', '0', '', '', 0, '2000-01-01', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `quest`
--

CREATE TABLE IF NOT EXISTS `quest` (
  `user_id` int(11) NOT NULL,
  `number` int(11) NOT NULL,
  `quest_id` int(11) NOT NULL,
  `npcs` varchar(64) NOT NULL,
  `npcstarget` varchar(64) NOT NULL,
  PRIMARY KEY (`user_id`,`number`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Volcado de datos para la tabla `quest`
--

INSERT INTO `quest` (`user_id`, `number`, `quest_id`, `npcs`, `npcstarget`) VALUES
(2355, 1, 4, '10-0-15-5', ''),
(2355, 2, 0, '', ''),
(2355, 3, 0, '', ''),
(2355, 4, 0, '', ''),
(2355, 5, 0, '', ''),
(2357, 1, 0, '', ''),
(2357, 2, 0, '', ''),
(2357, 3, 0, '', ''),
(2357, 4, 0, '', ''),
(2357, 5, 0, '', ''),
(2358, 1, 0, '', ''),
(2358, 2, 0, '', ''),
(2358, 3, 0, '', ''),
(2358, 4, 0, '', ''),
(2358, 5, 0, '', ''),
(2360, 1, 0, '', ''),
(2360, 2, 0, '', ''),
(2360, 3, 0, '', ''),
(2360, 4, 0, '', ''),
(2360, 5, 0, '', ''),
(2362, 1, 3, '', ''),
(2362, 2, 0, '', ''),
(2362, 3, 0, '', ''),
(2362, 4, 0, '', ''),
(2362, 5, 0, '', ''),
(2364, 1, 2, '', ''),
(2364, 2, 0, '', ''),
(2364, 3, 0, '', ''),
(2364, 4, 0, '', ''),
(2364, 5, 0, '', ''),
(2365, 1, 0, '', ''),
(2365, 2, 0, '', ''),
(2365, 3, 0, '', ''),
(2365, 4, 0, '', ''),
(2365, 5, 0, '', ''),
(2366, 1, 0, '', ''),
(2366, 2, 0, '', ''),
(2366, 3, 0, '', ''),
(2366, 4, 0, '', ''),
(2366, 5, 0, '', ''),
(2367, 1, 0, '', ''),
(2367, 2, 0, '', ''),
(2367, 3, 0, '', ''),
(2367, 4, 0, '', ''),
(2367, 5, 0, '', ''),
(2368, 1, 0, '', ''),
(2368, 2, 0, '', ''),
(2368, 3, 0, '', ''),
(2368, 4, 0, '', ''),
(2368, 5, 0, '', ''),
(2369, 1, 0, '', ''),
(2369, 2, 0, '', ''),
(2369, 3, 0, '', ''),
(2369, 4, 0, '', ''),
(2369, 5, 0, '', ''),
(2370, 1, 0, '', ''),
(2370, 2, 0, '', ''),
(2370, 3, 0, '', ''),
(2370, 4, 0, '', ''),
(2370, 5, 0, '', ''),
(2371, 1, 0, '', ''),
(2371, 2, 0, '', ''),
(2371, 3, 0, '', ''),
(2371, 4, 0, '', ''),
(2371, 5, 0, '', ''),
(2372, 1, 0, '', ''),
(2372, 2, 0, '', ''),
(2372, 3, 0, '', ''),
(2372, 4, 0, '', ''),
(2372, 5, 0, '', ''),
(2373, 1, 0, '', ''),
(2373, 2, 0, '', ''),
(2373, 3, 0, '', ''),
(2373, 4, 0, '', ''),
(2373, 5, 0, '', ''),
(2374, 1, 0, '', ''),
(2374, 2, 0, '', ''),
(2374, 3, 0, '', ''),
(2374, 4, 0, '', ''),
(2374, 5, 0, '', ''),
(2375, 1, 0, '', ''),
(2375, 2, 0, '', ''),
(2375, 3, 0, '', ''),
(2375, 4, 0, '', ''),
(2375, 5, 0, '', ''),
(2376, 1, 0, '', ''),
(2376, 2, 0, '', ''),
(2376, 3, 0, '', ''),
(2376, 4, 0, '', ''),
(2376, 5, 0, '', ''),
(2377, 1, 1, '', ''),
(2377, 2, 3, '0-0-0-0', ''),
(2377, 3, 0, '', ''),
(2377, 4, 0, '', ''),
(2377, 5, 0, '', ''),
(2378, 1, 0, '', ''),
(2378, 2, 0, '', ''),
(2378, 3, 0, '', ''),
(2378, 4, 0, '', ''),
(2378, 5, 0, '', ''),
(2379, 1, 0, '', ''),
(2379, 2, 0, '', ''),
(2379, 3, 0, '', ''),
(2379, 4, 0, '', ''),
(2379, 5, 0, '', ''),
(2380, 1, 0, '', ''),
(2380, 2, 0, '', ''),
(2380, 3, 0, '', ''),
(2380, 4, 0, '', ''),
(2380, 5, 0, '', '');

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `quest_done`
--

CREATE TABLE IF NOT EXISTS `quest_done` (
  `user_id` mediumint(8) unsigned NOT NULL,
  `quest_id` tinyint(3) unsigned NOT NULL,
  PRIMARY KEY (`user_id`,`quest_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Volcado de datos para la tabla `quest_done`
--

INSERT INTO `quest_done` (`user_id`, `quest_id`) VALUES
(2355, 0),
(2355, 2),
(2355, 3),
(2355, 5),
(2357, 0),
(2358, 0),
(2360, 0),
(2362, 2),
(2364, 0),
(2365, 0),
(2366, 0),
(2367, 0),
(2368, 0),
(2369, 0),
(2370, 0),
(2371, 0),
(2372, 0),
(2373, 0),
(2374, 0),
(2375, 0),
(2376, 0),
(2377, 0),
(2378, 0),
(2379, 0),
(2380, 0);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `vault`
--

CREATE TABLE IF NOT EXISTS `vault` (
  `cuenta_id` bigint(20) NOT NULL,
  `slot` int(11) NOT NULL,
  `item` int(11) NOT NULL DEFAULT '0',
  `quantity` int(11) NOT NULL DEFAULT '0',
  `updated_at` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  PRIMARY KEY (`cuenta_id`,`slot`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Volcado de datos para la tabla `vault`
--

INSERT INTO `vault` (`cuenta_id`, `slot`, `item`, `quantity`, `updated_at`) VALUES
(1325, 1, 402, 1, '0000-00-00 00:00:00');

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;

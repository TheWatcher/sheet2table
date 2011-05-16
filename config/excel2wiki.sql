-- phpMyAdmin SQL Dump
-- version 3.3.6
-- http://www.phpmyadmin.net
--
-- Host: localhost
-- Generation Time: May 16, 2011 at 03:30 PM
-- Server version: 5.1.56
-- PHP Version: 5.3.6-pl0-gentoo

SET SQL_MODE="NO_AUTO_VALUE_ON_ZERO";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

--
-- Database: `excel2wiki`
--

-- --------------------------------------------------------

--
-- Table structure for table `e2t_formats`
--

CREATE TABLE IF NOT EXISTS `e2t_formats` (
  `id` tinyint(3) unsigned NOT NULL AUTO_INCREMENT,
  `name` varchar(32) NOT NULL COMMENT 'The name of the output format',
  `function` varchar(255) NOT NULL COMMENT 'The output formatter function',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM  DEFAULT CHARSET=utf8 AUTO_INCREMENT=3 ;

--
-- Dumping data for table `e2t_formats`
--

INSERT INTO `e2t_formats` (`id`, `name`, `function`) VALUES
(0, 'Mediawiki table markup', 'format_mediawiki'),
(1, 'HTML table markup', 'format_html');

-- --------------------------------------------------------

--
-- Table structure for table `e2t_headers`
--

CREATE TABLE IF NOT EXISTS `e2t_headers` (
  `sheetid` mediumint(8) unsigned NOT NULL COMMENT 'The id of the sheet in e2t_sheets this is a header cell for',
  `colnum` mediumint(8) unsigned NOT NULL COMMENT 'The column the header is in',
  `rownum` mediumint(8) unsigned NOT NULL COMMENT 'The row the header is in',
  UNIQUE KEY `header` (`sheetid`,`colnum`,`rownum`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

--
-- Dumping data for table `e2t_headers`
--


-- --------------------------------------------------------

--
-- Table structure for table `e2t_popups`
--

CREATE TABLE IF NOT EXISTS `e2t_popups` (
  `sheetid` mediumint(8) unsigned NOT NULL COMMENT 'The id of the sheet in e2t_sheets this belongs to',
  `popupid` mediumint(8) unsigned NOT NULL COMMENT 'The id of the popup on the sheet',
  `title_col` mediumint(8) unsigned NOT NULL COMMENT 'The id of the column that should form the popup anchor',
  `body_col` mediumint(8) unsigned NOT NULL COMMENT 'The id of the column used to make popup bodies.'
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

--
-- Dumping data for table `e2t_popups`
--


-- --------------------------------------------------------

--
-- Table structure for table `e2t_settings`
--

CREATE TABLE IF NOT EXISTS `e2t_settings` (
  `name` varchar(255) NOT NULL,
  `value` varchar(255) NOT NULL,
  PRIMARY KEY (`name`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 COMMENT='Site settings';

--
-- Dumping data for table `e2t_settings`
--

INSERT INTO `e2t_settings` (`name`, `value`) VALUES
('base', '/path/to/excel2table'),
('scriptpath', '/'),
('default_style', 'default'),
('force_style', '1'),
('logfile', ''),
('site_url', 'http://url.of/excel2table'),
('file_dir', '/path/to/excel2table_data'),
('ip_security', '3'),
('last_gc', '1305552338');

-- --------------------------------------------------------

--
-- Table structure for table `e2t_sheets`
--

CREATE TABLE IF NOT EXISTS `e2t_sheets` (
  `id` mediumint(8) unsigned NOT NULL AUTO_INCREMENT,
  `source_name` varchar(255) NOT NULL COMMENT 'The name the file was originally uploaded with',
  `local_name` varchar(64) NOT NULL COMMENT 'The name the sheet is stored as in the file directory',
  `file_type` enum('xls','xlsx') NOT NULL COMMENT 'The type of the excel file (old or new)',
  `sheet_num` mediumint(8) unsigned DEFAULT NULL COMMENT 'The selected worksheet, NULL if not set',
  `output_type` tinyint(3) unsigned DEFAULT NULL COMMENT 'The requested output type',
  `zebra` tinyint(3) unsigned DEFAULT NULL COMMENT 'Should zebra table class be added?',
  `set_headers` tinyint(1) unsigned DEFAULT NULL COMMENT 'If true, the user has opted to set headers for the table',
  `set_popups` tinyint(1) unsigned DEFAULT NULL COMMENT 'If true the user has opted to set popups in the table',
  `remote_addr` char(15) NOT NULL COMMENT 'The IP address of the host doing the request',
  `last_updated` int(10) unsigned NOT NULL COMMENT 'Last modified timestamp',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM  DEFAULT CHARSET=utf8 AUTO_INCREMENT=18 ;

--
-- Dumping data for table `e2t_sheets`
--


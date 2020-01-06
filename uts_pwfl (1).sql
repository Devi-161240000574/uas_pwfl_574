-- phpMyAdmin SQL Dump
-- version 4.5.1
-- http://www.phpmyadmin.net
--
-- Host: 127.0.0.1
-- Generation Time: Jan 06, 2020 at 12:36 PM
-- Server version: 10.1.19-MariaDB
-- PHP Version: 5.6.28

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `uts_pwfl`
--

-- --------------------------------------------------------

--
-- Table structure for table `histori`
--

CREATE TABLE `histori` (
  `idmhs` int(5) NOT NULL,
  `nim` varchar(15) NOT NULL,
  `no_hp_lama` varchar(15) NOT NULL,
  `no_hp_baru` varchar(15) NOT NULL,
  `tgl_diubah` date NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `histori`
--

INSERT INTO `histori` (`idmhs`, `nim`, `no_hp_lama`, `no_hp_baru`, `tgl_diubah`) VALUES
(1, '161240000574', '081326799926', '081326799928', '2019-10-29'),
(2, '161240000677', '085123451765', '085123451777', '2019-11-11'),
(3, '161240000677', '085123451777', '085123451777', '2019-11-11'),
(4, '161240000677', '085123451777', '085123451777', '2019-11-11');

-- --------------------------------------------------------

--
-- Table structure for table `mhs`
--

CREATE TABLE `mhs` (
  `idmhs` int(5) NOT NULL,
  `nim` varchar(12) NOT NULL,
  `nama` varchar(40) NOT NULL,
  `jenis_kelamin` varchar(10) NOT NULL,
  `alamat` varchar(200) NOT NULL,
  `nomer_hp` varchar(15) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mhs`
--

INSERT INTO `mhs` (`idmhs`, `nim`, `nama`, `jenis_kelamin`, `alamat`, `nomer_hp`) VALUES
(1, '161240000574', 'Devi Nurlita Andriyani', 'Perempuan', 'Sinanggul Sidang', '081326799928'),
(2, '161240000580', 'Melati Nur Indah Sari', 'Perempuan', 'Ngabul', '085651888275'),
(3, '161240000583', 'Heni Mundjayati', 'Perempuan', 'Nalumsari', '089255417333'),
(4, '161240000677', 'Mahardika Widi', 'Laki-laki', 'Jepara', '085123451777'),
(10, '161240000555', 'Andriyan', 'Laki-Laki', 'Pati', '0851234511111'),
(11, '1612400005', 'Cahya', 'Laki-Laki', 'Jepara', '089255417336'),
(12, '15146776', 'Ani', 'Perempuan', 'Jepara', '08166222671'),
(13, '161240000570', 'Sino', 'Laki-Laki', 'Jepara', '08166222673'),
(14, '161240000559', 'Hinoa', 'Laki-Laki', 'Pati', '089255417336'),
(15, '161240000559', 'Cahyo', 'Laki-Laki', 'Jepara', '0851234511112'),
(16, '16124000567', 'Minan', 'Laki-Laki', 'Pati', '089255417334');

--
-- Triggers `mhs`
--
DELIMITER $$
CREATE TRIGGER `membuat_log` BEFORE UPDATE ON `mhs` FOR EACH ROW BEGIN
    INSERT INTO histori
    SET nim=old.nim,
    no_hp_lama=old.nomer_hp,
    no_hp_baru=new.nomer_hp,
    tgl_diubah=NOW() ;
END
$$
DELIMITER ;

-- --------------------------------------------------------

--
-- Stand-in structure for view `tampil`
--
CREATE TABLE `tampil` (
`idmhs` int(5)
,`nim` varchar(12)
,`nama` varchar(40)
,`jenis_kelamin` varchar(10)
,`alamat` varchar(200)
,`nomer_hp` varchar(15)
);

-- --------------------------------------------------------

--
-- Structure for view `tampil`
--
DROP TABLE IF EXISTS `tampil`;

CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `tampil`  AS  (select `mhs`.`idmhs` AS `idmhs`,`mhs`.`nim` AS `nim`,`mhs`.`nama` AS `nama`,`mhs`.`jenis_kelamin` AS `jenis_kelamin`,`mhs`.`alamat` AS `alamat`,`mhs`.`nomer_hp` AS `nomer_hp` from `mhs`) ;

--
-- Indexes for dumped tables
--

--
-- Indexes for table `histori`
--
ALTER TABLE `histori`
  ADD PRIMARY KEY (`idmhs`);

--
-- Indexes for table `mhs`
--
ALTER TABLE `mhs`
  ADD PRIMARY KEY (`idmhs`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `histori`
--
ALTER TABLE `histori`
  MODIFY `idmhs` int(5) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=5;
--
-- AUTO_INCREMENT for table `mhs`
--
ALTER TABLE `mhs`
  MODIFY `idmhs` int(5) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=17;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;

# Host: localhost  (Version: 5.5.5-10.1.34-MariaDB)
# Date: 2020-06-29 18:14:42
# Generator: MySQL-Front 5.3  (Build 4.263)

/*!40101 SET NAMES latin1 */;

#
# Structure for table "tb_daftar"
#

DROP TABLE IF EXISTS `tb_daftar`;
CREATE TABLE `tb_daftar` (
  `id_daftar` char(7) NOT NULL DEFAULT '',
  `tgl_daftar` date DEFAULT NULL,
  `kelas` varchar(10) DEFAULT NULL,
  `tahun` char(4) DEFAULT NULL,
  `nis` char(8) DEFAULT NULL,
  `pendapatan` mediumint(8) DEFAULT NULL,
  `nilai` decimal(4,2) DEFAULT NULL,
  `saudara` tinyint(3) DEFAULT NULL,
  PRIMARY KEY (`id_daftar`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Data for table "tb_daftar"
#

INSERT INTO `tb_daftar` VALUES ('p0001','2020-06-28','Kelas2','2018','19123415',2000000,90.00,4),('p0002','2020-06-28','Kelas1','2017','19123416',3000000,95.00,2),('p0003','2020-06-29','Kelas2','2018','01912341',8388607,98.00,2);

#
# Structure for table "tb_login"
#

DROP TABLE IF EXISTS `tb_login`;
CREATE TABLE `tb_login` (
  `username` varchar(100) NOT NULL DEFAULT '',
  `pasword` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`username`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Data for table "tb_login"
#

INSERT INTO `tb_login` VALUES ('kharisma','imma');

#
# Structure for table "tb_nilai"
#

DROP TABLE IF EXISTS `tb_nilai`;
CREATE TABLE `tb_nilai` (
  `id_nilai` char(5) NOT NULL,
  `id_daftar` char(5) DEFAULT NULL,
  `n_pendapatan` decimal(4,2) DEFAULT NULL,
  `n_nilai` decimal(4,2) DEFAULT NULL,
  `n_saudara` decimal(4,2) DEFAULT NULL,
  `preferensi` decimal(4,2) DEFAULT NULL,
  PRIMARY KEY (`id_nilai`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Data for table "tb_nilai"
#


#
# Structure for table "tb_siswa"
#

DROP TABLE IF EXISTS `tb_siswa`;
CREATE TABLE `tb_siswa` (
  `nis` char(8) NOT NULL DEFAULT '',
  `nama` varchar(40) DEFAULT NULL,
  `t_lahir` varchar(25) DEFAULT NULL,
  `tgl_lahir` varchar(30) DEFAULT NULL,
  `jk` char(9) DEFAULT NULL,
  `alamat` varchar(100) DEFAULT NULL,
  `jurusan` varchar(20) DEFAULT NULL,
  `telpon` varchar(15) DEFAULT NULL,
  PRIMARY KEY (`nis`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Data for table "tb_siswa"
#

INSERT INTO `tb_siswa` VALUES ('01912341','silva','sewo','12 januari','laki-laki','sewo','BAHASA','082343195774'),('09123419','arni','batu-batu','2 februari 1999','perempuan','batu-batu','IPA','082234582213'),('19123415','kharisma nur','lajarella','23 Maret 1998','perempuan','lajarella','IPS','085242195724'),('19123416','fenny Azis','leworeng','05 juni 1997','perempuan','leworeng','TKJ','082234562343'),('19123417','akbar','soppeng','05 maret 1998','laki-laki','soppeng','IPA','085464192456');

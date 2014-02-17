-- phpMyAdmin SQL Dump
-- version 3.5.2.2
-- http://www.phpmyadmin.net
--
-- Host: 127.0.0.1
-- Generation Time: Feb 13, 2014 at 01:30 PM
-- Server version: 5.5.27
-- PHP Version: 5.4.7

SET SQL_MODE="NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

--
-- Database: `invent_db`
--

DELIMITER $$
--
-- Procedures
--
CREATE DEFINER=`root`@`localhost` PROCEDURE `AmbilKirim`(IN `kode` VARCHAR(15))
BEGIN
START TRANSACTION;
	SELECT tbkirimun.*,tbdetkirimun.*,tbunit.namaunit,tbbarang.namaBarang
FROM tbkirimun,tbdetkirimun,tbunit,tbbarang
WHERE tbkirimun.kdKirimun=tbdetkirimun.kdKirimun
AND tbkirimun.kdunit=tbunit.kdunit
AND tbdetkirimun.kdBarang=tbbarang.kdBarang
and tbkirimun.kdKirimun=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `AmbilTerima`(kode VARCHAR(15))
BEGIN
START TRANSACTION;
	SELECT tbterima.*,tbdetterima.*,tbdistributor.namaDistributor,tbbarang.namaBarang
FROM tbterima,tbdetterima,tbdistributor,tbbarang
WHERE tbterima.kdKirim=tbdetterima.kdKirim
AND tbterima.kdDistributor=tbdistributor.kdDistributor
AND tbdetterima.kdBarang=tbbarang.kdBarang
and tbterima.kdKirim=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `BarangJenis`(IN `kode` VARCHAR(20))
BEGIN
START TRANSACTION;
SELECT tbjenis.nama as jenis,tbmerk.nama as merk,
tbbarang.* 
from tbbarang,tbjenis,tbmerk
 WHERE tbbarang.idjenis=tbjenis.idjenis
AND tbbarang.idmerk=tbmerk.idmerk
AND tbjenis.nama LIKE kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `BarangKode`(IN `kode` VARCHAR(20))
BEGIN
START TRANSACTION;
SELECT tbjenis.nama as jenis,tbmerk.nama as merk,
tbbarang.* 
from tbbarang,tbjenis,tbmerk
 WHERE tbbarang.idjenis=tbjenis.idjenis
AND tbbarang.idmerk=tbmerk.idmerk
AND tbbarang.kdBarang LIKE kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `BarangMerk`(IN `kode` VARCHAR(20))
BEGIN
START TRANSACTION;
SELECT tbjenis.nama as jenis,tbmerk.nama as merk,
tbbarang.* 
from tbbarang,tbjenis,tbmerk
 WHERE tbbarang.idjenis=tbjenis.idjenis
AND tbbarang.idmerk=tbmerk.idmerk
AND tbmerk.nama LIKE kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `BarangNama`(IN `kode` VARCHAR(20))
BEGIN
START TRANSACTION;
SELECT tbjenis.nama as jenis,tbmerk.nama as merk,
tbbarang.* 
from tbbarang,tbjenis,tbmerk
 WHERE tbbarang.idjenis=tbjenis.idjenis
AND tbbarang.idmerk=tbmerk.idmerk
AND tbbarang.namaBarang LIKE kode order by tbbarang.namabarang;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `CekBarang`(IN `kode` VARCHAR(20))
BEGIN
START TRANSACTION;
SELECT kdbarang FROM tbstok WHERE kdbarang=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `CekDist`(IN `kode` VARCHAR(5))
BEGIN
START TRANSACTION;
SELECT kdDistributor 
FROM tbterima where kdDistributor=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `CekJenis`(kode VARCHAR(5))
BEGIN
START TRANSACTION;
SELECT idjenis FROM tbbarang where idjenis=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `CekMerk`(IN `kode` VARCHAR(5))
BEGIN
START TRANSACTION;
select idMERK from tbbarang where idmerk=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `CekUnit`(IN `kode` VARCHAR(15))
BEGIN
START TRANSACTION;
SELECT kdunit FROM 
tbkirimun where kdkirimun=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `EditBarang`(
kode CHAR(15),nama CHAR(25), satuan CHAR(10), idjenis CHAR(3), idmerk CHAR(3),ket TEXT,user CHAR(20), tgl DATETIME)
BEGIN
START TRANSACTION;
UPDATE tbbarang SET namaBarang=nama,
satuan=satuan,
idmerk=idmerk,
idjenis=idjenis,
keterangan=ket,
user_ubah=user,
tgl_ubah=tgl WHERE kdBarang=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `EditDist`(IN `kode` VARCHAR(5), IN `nama` VARCHAR(25), IN `alamat` TEXT, IN `telp` INT(13), IN `kontak` VARCHAR(25), IN `user` VARCHAR(20), IN `tgl` DATETIME)
BEGIN
START TRANSACTION;
UPDATE tbdistributor SET namaDistributor=nama,
alamat=alamat,
telp=telp,
kontakPerson=kontak,
user_ubah=user,
tgl_ubah=tgl
where kdDistributor=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `EditJenis`(kode VARCHAR(5),nama varchar(20),user varchar(20),tgl DATETIME)
BEGIN
START TRANSACTION;
UPDATE tbjenis SET nama=nama,user_ubah=user,tgl_ubah=tgl WHERE idjenis=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `EditKirimUn`(IN `kode` VARCHAR(15), IN `tgl` DATE, IN `kodeu` VARCHAR(15), IN `userubah` VARCHAR(20), IN `tglubah` DATE, IN `konf` VARCHAR(1))
BEGIN
START TRANSACTION;
	UPDATE tbKirimUn SET tglKirim=tgl,kdUnit=kodeu,user_ubah=userubah,tgl_ubah=tglubah,konfirm=konf
	WHERE kdkirimun=kode;
	DELETE FROM tbdetkirimun WHERE kdKirimun=kode;
	INSERT INTO tbdetkirimun(kdKirimUn,kdBarang,jml,harga,total) 
	SELECT tbtmpkirimun.kdKirimUn,tbtmpkirimun.kdbarang,tbtmpkirimun.jml,tbtmpkirimun.harga,tbtmpkirimun.total
	from tbtmpkirimun WHERE tbtmpkirimun.kdKirimun=kode;
	UPDATE tbtmpkirimun set kdKirimUn=kode where kdKirimUn is null;
	
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `EditMerk`(kode CHAR(15),nama CHAR(25), user CHAR(20), tgl DATETIME)
BEGIN
START TRANSACTION;
UPDATE tbmerk SET nama=nama,user_ubah=user,tgl_ubah=tgl WHERE idmerk=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `EditRetsup`(IN `kode` VARCHAR(16), IN `tgl` DATE, IN `kodek` VARCHAR(16), IN `userubah` VARCHAR(20), IN `tglubah` DATETIME, IN `konf` VARCHAR(1))
BEGIN
START TRANSACTION;
	UPDATE tbRetursup set tglretur=tgl,kdKirim=kodek,user_ubah=userubah,tgl_ubah=tglubah,konfirm=konf
	WHERE kdretursup=kode;
	
	DELETE FROM tbdetretursup where kdretursup=kode;
	INSERT INTO tbdetretursup(kdReturSup,kdBarang,jml,harga,total,alasan) 
	SELECT tbtmpretursup.kdReturSup,tbtmpretursup.kdbarang,tbtmpretursup.jml,tbtmpretursup.harga,tbtmpretursup.total,tbtmpretursup.alasan 
	from tbtmpretursup WHERE tbtmpretursup.kdReturSup=kode;
	UPDATE tbdetretursup set kdReturSup=kode where kdReturSup is null;
	
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `EditRetUn`(IN `kode` VARCHAR(16), IN `tgl` DATE, IN `kodek` VARCHAR(16), IN `userubah` VARCHAR(20), IN `tglubah` DATETIME, IN `konf` VARCHAR(1))
BEGIN
START TRANSACTION;
	UPDATE tbReturUn set tglreturUn=tgl,kdKirimUn=kodek,user_ubah=userubah,tgl_ubah=tglubah,konfirm=konf
	WHERE kdreturun=kode;
	
	DELETE FROM tbdetreturun where kdreturun=kode;
	INSERT INTO tbdetreturun(kdReturun,kdBarang,jml,harga,total,alasan) 
	SELECT tbtmpreturun.kdReturun,tbtmpreturun.kdbarang,tbtmpreturun.jml,tbtmpreturun.harga,tbtmpreturun.total,tbtmpreturun.alasan 
	from tbtmpreturun WHERE tbtmpreturun.kdReturun=kode;
	UPDATE tbdetreturun set kdReturUn=kode where kdReturUn is null;
	
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `EditTerima`(IN `kode` VARCHAR(15), IN `tglTerima` DATE, IN `nofaktur` VARCHAR(20), IN `tglf` DATE, IN `kdDis` VARCHAR(12), IN `useru` VARCHAR(25), IN `tglUbah` DATETIME)
BEGIN
START TRANSACTION;
	UPDATE tbtmpterima1 SET tgl_terima=tglterima,nofaktur=nofaktur,tglfaktur=tglf,
			kddistributor=kddis,user_ubah=useru,tgl_ubah=tglubah WHERE Kdkirim=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `EditTmpKirUn`(kode VARCHAR(5),Kodeb varchar(15),jml INT(12),
hrg INT(13),tot INT(15))
BEGIN
START TRANSACTION;
UPDATE tbtmpkirimun SET jml=jml,harga=hrg,total=tot WHERE
kdkirimun=kode and kdbarang=kodeb;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `EditTmpRetSup`(kode VARCHAR(15),kodeB varchar(15),
								qty INT(12),harga INT(15),total INT(15),alasan TEXT,kodek varchar(15))
BEGIN
START TRANSACTION;
update tbtmpRetursup SET jml=qty,harga=harga,total=total,
alasan=alasan
where  kdBarang=kodeB and kdretursup=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `EditTmpRetUn`(kode VARCHAR(15),kodeB varchar(15),
								qty INT(12),harga INT(15),total INT(15),alasan TEXT)
BEGIN
START TRANSACTION;
update tbtmpReturUn SET jml=qty,harga=harga,total=total,
alasan=alasan
where  kdBarang=kodeB and kdreturUn=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `EditTmpTer`(IN `kode` VARCHAR(15), IN `kodeB` VARCHAR(15), IN `qty` INT(12), IN `hargaD` INT(15), IN `jumlahD` INT(15), IN `hargaF` INT(15), IN `persen` INT(3))
BEGIN
START TRANSACTION;
update tbtmpterima SET jumlah=qty,hargaDasar=hargaD,totalDasar=jumlahD,
hargaFixed=hargaF,persen=persen
where kdKirim=kode and kdBarang=kodeB;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `EditUnit`(kode VARCHAR(15),nama varchar(20),alamat TEXT,telp INT(13),kontak VARCHAR(25),
user varchar(20),tgl DATETIME)
BEGIN
START TRANSACTION;
UPDATE tbunit SET namaunit=nama,alamat=alamat,
telp=telp,kontakperson=kontak,user_ubah=user,tgl_ubah=tgl WHERE kdunit=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `EditUser`(IN `nama` VARCHAR(30), IN `usern` VARCHAR(20), IN `pass` TEXT, IN `lev` VARCHAR(13), IN `user` VARCHAR(20), IN `tgl` DATETIME, IN `stat` VARCHAR(1))
BEGIN
START TRANSACTION;
UPDATE tbuser SET password=pass,user_ubah=user,tgl_ubah=tgl,status=stat,level=lev WHERE iduser=usern and namalengkap=nama;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `HapusBarang`(IN `kode` VARCHAR(5))
BEGIN
START TRANSACTION;
DELETE FROM tbbarang WHERE kdbarang=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `HapusDist`(kode VARCHAR(5))
BEGIN
START TRANSACTION;
DELETE FROM tbdistributor WHERE kdDistributor=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `HapusJenis`(kode VARCHAR(5))
BEGIN
START TRANSACTION;
DELETE FROM tbjenis WHERE idjenis=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `HapusMerk`(kode VARCHAR(5))
BEGIN
START TRANSACTION;
DELETE FROM tbmerk WHERE idmerk=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `HapusTmpKirUn`(kode VARCHAR(15),kodeb VARCHAR(15))
BEGIN
START TRANSACTION;
DELETE FROM tbtmpkirimun WHERE kdkirimun=kode and kdbarang=kodeb;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `HapusTmpRetsup`(kode VARCHAR(15),kodeb VARCHAR(15))
BEGIN
START TRANSACTION;
DELETE FROM tbtmpretursup WHERE kdretursup=kode and kdbarang=kodeb;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `HapusTmpRetUn`(kode VARCHAR(15),kodeb VARCHAR(15))
BEGIN
START TRANSACTION;
DELETE FROM tbtmpreturUn WHERE kdreturUn=kode and kdbarang=kodeb;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `HapusTmpTer`(IN `kode` VARCHAR(15), IN `kodeB` VARCHAR(15))
BEGIN
START TRANSACTION;
DELETE FROM tbtmpterima where kdKirim=kode and kdBarang=kodeB;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `HapusUnit`(kode VARCHAR(15))
BEGIN
START TRANSACTION;
DELETE FROM tbunit WHERE kdunit=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `HapusUser`(kode VARCHAR(25))
BEGIN
START TRANSACTION;
DELETE FROM tbuser WHERE username=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `KirimKode`(IN `kode` VARCHAR(15))
BEGIN
START TRANSACTION;
SELECT tbtmpkirimun1.KdkirimUn,tbtmpkirimun1.tglkirim,tbtmpkirimun1.user_ubah,tbtmpkirimun1.tgl_ubah,
tbtmpkirimun1.kdunit,tbunit.namaunit 
FROM tbtmpkirimun1,TBTMPKIRIMUN,tbunit
WHERE tbtmpkirimun1.kdkirimUn=TBTMPKIRIMUN.kdkirimUn
AND tbUnit.kdUnit=tbtmpkirimun1.kdUnit
AND tbtmpkirimun1.kdkirimUn LIKE kode 
and tbtmpkirimun1.flag='N'
GROUP BY tbtmpkirimun1.kdkirimUn;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `KirimUnit`(IN `kode` VARCHAR(15))
BEGIN
START TRANSACTION;
SELECT tbtmpkirimun1.KdkirimUn,tbtmpkirimun1.tglkirim,tbtmpkirimun1.user_ubah,tbtmpkirimun1.tgl_ubah,
tbtmpkirimun1.kdunit,tbunit.namaunit 
FROM tbtmpkirimun1,TBTMPKIRIMUN,tbunit
WHERE tbtmpkirimun1.kdkirimUn=TBTMPKIRIMUN.kdkirimUn
AND tbUnit.kdUnit=tbtmpkirimun1.kdUnit
AND tbtmpkirimUn1.kdUnit LIKE kode
and tbtmpkirimun1.flag='N'
GROUP BY tbtmpkirimun1.kdkirimUn;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `retur`(IN `kode` VARCHAR(15))
BEGIN
START TRANSACTION;
select tbretursup.*,tbbarang.namaBarang,tbdetretursup.*,tbdistributor.namadistributor,tbdistributor.kddistributor,
tbterima.tgl_terima,tbterima.nofaktur,tbterima.tglfaktur
from tbretursup,tbdetretursup,tbdistributor,tbterima,tbbarang
where tbretursup.kdkirim=tbterima.kdkirim
and tbretursup.kdretursup=tbdetretursup.kdretursup
and tbterima.kddistributor=tbdistributor.kddistributor
and tbdetretursup.kdbarang=tbbarang.kdbarang
AND tbretursup.kdretursup=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `rETURDis`(IN `kode` VARCHAR(15))
BEGIN
START TRANSACTION;
select tbtmpretursup1.*,tbdistributor.namadistributor
from tbtmpretursup1,tbdistributor
where  tbtmpretursup1.kddistributor=tbdistributor.kddistributor
AND tbdistributor.kddistributor=kode
and tbtmpretursup1.flag='N'
group by tbtmpretursup1.kdretursup;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `rETURKode`(IN `kode` VARCHAR(15))
BEGIN
START TRANSACTION;
select tbtmpretursup1.*,tbdistributor.namadistributor
from tbtmpretursup1,tbdistributor
where  tbtmpretursup1.kddistributor=tbdistributor.kddistributor
AND tbtmpretursup1.kdretursup=kode
and tbtmpretursup1.flag='N'
group by tbtmpretursup1.kdretursup;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `returU`(IN `kode` VARCHAR(15))
BEGIN
START TRANSACTION;
select tbreturun.*,tbbarang.namaBarang,tbdetreturun.*,tbunit.namaunit,tbunit.kdunit,
tbkirimun.tglkirim
from tbreturun,tbdetreturun,tbunit,tbkirimun,tbbarang
where tbreturUn.kdkirimUn=tbKirimUn.kdkirimUn
and tbreturUn.kdreturUn=tbdetreturUn.kdreturUn
and tbKirimUn.kdUnit=tbUnit.kdUnit
and tbdetreturun.kdbarang=tbbarang.kdbarang
AND tbreturun.kdreturun=kode;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `ReturUKode`(IN `kode` VARCHAR(15))
BEGIN
START TRANSACTION;
select tbtmpreturun1.*,tbunit.namaunit
from tbtmpreturun1,tbUnit
where tbtmpreturun1.kdUnit=tbUnit.kdUnit
and tbtmpreturun1.flag='N'
AND tbtmpreturun1.kdreturUn LIKE kode
group by tbtmpreturun1.kdreturUn;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `ReturUnit`(IN `kode` VARCHAR(15))
BEGIN
START TRANSACTION;
select tbtmpreturun1.*,tbunit.namaunit
from tbtmpreturun1,tbUnit
where tbtmpreturun1.kdUnit=tbUnit.kdUnit
and tbtmpreturun1.flag='N'
AND tbkirimun.kdkirimUn LIKE kode
group by tbtmpreturun1.kdreturUn;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahBarang`(IN `kode` CHAR(15), IN `nama` CHAR(25), IN `satuan` CHAR(10), IN `idjenis` CHAR(12), IN `idmerk` CHAR(12), IN `ket` TEXT, IN `user` CHAR(20), IN `tgl` DATETIME, IN `masuk` INT(12), IN `keluar` INT(12), IN `stok` INT(12))
BEGIN
START TRANSACTION;
INSERT INTO tbbarang(kdBarang,namaBarang,satuan,idmerk,idjenis,keterangan,user_ubah,tgl_ubah,stokAkhir)
VALUES(kode,nama,satuan,idjenis,idmerk,ket,user,tgl,stok);

INSERT INTO tbstok(kdBarang,no_bukti,masuk,keluar,stok,user_ubah,tgl_ubah,keterangan)
values(kode,'-',masuk,keluar,stok,user,tgl,'INPUT BARANG');
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahDist`(kode VARCHAR(12),nama varchar(25),alamat text,telp int(13),kontak varchar(25),user varchar(20),tgl DATETIME)
BEGIN
START TRANSACTION;
INSERT INTO tbdistributor(kdDistributor,namaDistributor,alamat,telp,kontakPerson,user_ubah,tgl_ubah)VALUES
(kode,nama,alamat,telp,kontak,user,tgl);
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahJenis`(kode VARCHAR(10),nama varchar(20),user varchar(20),tgl DATETIME)
BEGIN
START TRANSACTION;
INSERT INTO tbjenis(idjenis,nama,user_ubah,tgl_ubah)VALUES
(kode,nama,user,tgl);
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahKirimUn`(IN `kode` VARCHAR(15), IN `tgl` DATE, IN `kodeu` VARCHAR(15), IN `userubah` VARCHAR(20), IN `tglubah` DATE, IN `konf` VARCHAR(1))
BEGIN
START TRANSACTION;
	INSERT INTO tbKirimUn(kdKirimUn,tglKirim,kdUnit,user_ubah,tgl_ubah,konfirm)
	VALUES(kode,tgl,kodeu,userubah,tglubah,konf);
	
	INSERT INTO tbdetkirimun(kdKirimUn,kdBarang,jml,harga,total) 
	SELECT tbtmpkirimun.kdKirimUn,tbtmpkirimun.kdbarang,tbtmpkirimun.jml,tbtmpkirimun.harga,tbtmpkirimun.total
	from tbtmpkirimun WHERE tbtmpkirimun.kdKirimun=kode;
	UPDATE tbtmpkirimun set kdKirimUn=kode where kdKirimUn is null;
	
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahKirimUnSem`(IN `kode` VARCHAR(15), IN `tgl` DATE, IN `kodeu` VARCHAR(15), IN `userubah` VARCHAR(20), IN `tglubah` DATE)
BEGIN
START TRANSACTION;
	INSERT INTO tbTMPkirimun1(kdKirimUn,tglKirim,kdUnit,user_ubah,tgl_ubah)
	VALUES(kode,tgl,kodeu,userubah,tglubah);
	
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahMerk`(kode CHAR(15),nama CHAR(25), user CHAR(20), tgl DATETIME)
BEGIN
START TRANSACTION;
INSERT INTO tbmerk(idmerk,nama,user_ubah,tgl_ubah)VALUES(kode,nama,user,tgl);
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahRetsup`(IN `kode` VARCHAR(16), IN `tgl` DATE, IN `kodek` VARCHAR(16), IN `userubah` VARCHAR(20), IN `tglubah` DATETIME, IN `konf` VARCHAR(1))
BEGIN
START TRANSACTION;
	INSERT INTO tbRetursup(kdReturSup,tglretur,kddistributor,user_ubah,tgl_ubah,konfirm)
	VALUES(kode,tgl,kodek,userubah,tglubah,konf);
	
	INSERT INTO tbdetretursup(kdReturSup,kdBarang,jml,harga,total,alasan) 
	SELECT tbtmpretursup.kdReturSup,tbtmpretursup.kdbarang,tbtmpretursup.jml,tbtmpretursup.harga,tbtmpretursup.total,tbtmpretursup.alasan 
	from tbtmpretursup WHERE tbtmpretursup.kdReturSup=kode;
	UPDATE tbdetretursup set kdReturSup=kode where kdReturSup is null;
	
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahRetsupSem`(IN `kode` VARCHAR(16), IN `tgl` DATE, IN `kodek` VARCHAR(16), IN `userubah` VARCHAR(20), IN `tglubah` DATETIME)
BEGIN
START TRANSACTION;
	INSERT INTO tbtmpRetursup1(kdReturSup,tglretur,kdDISTRIBUTOR,user_ubah,tgl_ubah)
	VALUES(kode,tgl,kodek,userubah,tglubah);
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahRetUn`(IN `kode` VARCHAR(16), IN `tgl` DATE, IN `kodek` VARCHAR(16), IN `userubah` VARCHAR(20), IN `tglubah` DATETIME, IN `konf` VARCHAR(1))
BEGIN
START TRANSACTION;
	INSERT INTO tbReturUn(kdReturUn,tglreturUn,kdunit,user_ubah,tgl_ubah,konfirm)
	VALUES(kode,tgl,kodek,userubah,tglubah,konf);
	
	INSERT INTO tbdetreturUn(kdReturUn,kdBarang,jml,harga,total,alasan) 
	SELECT tbtmpreturUn.kdReturUn,tbtmpreturUn.kdbarang,tbtmpreturUn.jml,tbtmpreturUn.harga,tbtmpreturUn.total,tbtmpreturun.alasan 
	from tbtmpreturUn WHERE tbtmpreturUn.kdReturUn=kode;
	UPDATE tbdetreturUn set kdReturUn=kode where kdReturUn is null;
	
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahRetUnSem`(IN `kode` VARCHAR(16), IN `tgl` DATE, IN `kodek` VARCHAR(16), IN `userubah` VARCHAR(20), IN `tglubah` DATETIME, IN `konf` VARCHAR(1))
BEGIN
START TRANSACTION;
	INSERT INTO tbtmpReturUn1(kdReturUn,tglreturUn,kdUnit,user_ubah,tgl_ubah,flag)
	VALUES(kode,tgl,kodek,userubah,tglubah,konf);
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahTerima`(IN `kode` VARCHAR(15), IN `tglTerima` DATE, IN `nofaktur` VARCHAR(20), IN `Tgl_faktur` DATE, IN `kdDis` VARCHAR(12), IN `useru` VARCHAR(25), IN `tglUbah` DATETIME,`konfirm` VARCHAR(1))
BEGIN
START TRANSACTION;
	INSERT INTO tbterima(kdKirim,tgl_terima,nofaktur,tglfaktur,kdDistributor,user_ubah,tgl_ubah,konfirm)
	VALUES(kode,tglTerima,nofaktur,tgl_faktur,kdDis,useru,tglUbah,konfirm);
        INSERT INTO tbdetterima(kdBarang,jumlah,harga,total) SELECT tbtmpterima.kdbarang,tbtmpterima.jumlah,tbtmpterima.hargaDasar,tbtmpterima.totalDasar from tbtmpterima WHERE tbtmpterima.kdKirim=kode;
	UPDATE tbdetterima set kdKirim=kode where kdKirim is null;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahTerimaSem`(IN `kode` VARCHAR(15), IN `tglTerima` DATE, IN `nofaktur` VARCHAR(20), IN `Tgl_faktur` DATE, IN `kdDis` VARCHAR(12), IN `useru` VARCHAR(25), IN `tglUbah` DATETIME)
BEGIN
START TRANSACTION;
	INSERT INTO tbtmpterima1(kdKirim,tgl_terima,nofaktur,tglfaktur,kdDistributor,user_ubah,tgl_ubah)
	VALUES(kode,tglTerima,nofaktur,tgl_faktur,kdDis,useru,tglUbah);
	
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahTmpKir`(IN `kode` VARCHAR(15), IN `kodeB` VARCHAR(15), IN `jumlah` INT(12), IN `hargaD` INT(15), IN `totalD` INT(15), IN `hargaF` INT(15), IN `persen` INT(3))
BEGIN
START TRANSACTION;
INSERT INTO tbtmpterima(kdKirim,kdBarang,jumlah,hargaDasar,totalDasar,hargaFixed,persen)VALUES
(kode,kodeB,jumlah,hargaD,totalD,hargaF,persen);
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahTmpKirUn`(IN `kode` VARCHAR(15), IN `Kodeb` VARCHAR(15), IN `jml` INT(12), IN `hrg` INT(13), IN `tot` INT(15))
BEGIN
START TRANSACTION;
INSERT INTO tbtmpkirimun(kdkirimun,kdbarang,jml,harga,total)VALUES
(kode,Kodeb,jml,hrg,tot);
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahTmpRetSup`(IN `kode` VARCHAR(15), IN `kodeb` VARCHAR(15), IN `harga` INT(15), IN `jml` INT(15), IN `tot` INT(15), IN `ala` TEXT)
BEGIN
START TRANSACTION;
INSERT INTO tbtmpretursup(kdReturSup,kdBarang,harga,jml,total,alasan)VALUES
(kode,kodeb,harga,jml,tot,ala);
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahTmpRetUN`(IN `kode` VARCHAR(15), IN `kodeb` VARCHAR(15), IN `harga` INT(15), IN `jml` INT(15), IN `tot` INT(15), IN `ala` TEXT)
BEGIN
START TRANSACTION;
INSERT INTO tbtmpreturun(kdReturun,kdBarang,harga,jml,total,alasan)VALUES
(kode,kodeb,harga,jml,tot,ala);
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahUnit`(IN `kode` VARCHAR(15), IN `nama` VARCHAR(20), IN `alamat` TEXT, IN `telp` INT(13), IN `kontak` VARCHAR(25), IN `user` VARCHAR(20), IN `tgl` DATETIME)
BEGIN
START TRANSACTION;
INSERT INTO tbunit(kdunit,namaunit,alamat,telp,kontakperson,user_ubah,tgl_ubah)VALUES
(kode,nama,alamat,telp,kontak,user,tgl);
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TambahUser`(IN `nama` VARCHAR(30), IN `usern` VARCHAR(20), IN `pass` TEXT, IN `lev` VARCHAR(13), IN `user` VARCHAR(20), IN `tgl` DATETIME, IN `stat` VARCHAR(1))
BEGIN
START TRANSACTION;
INSERT INTO tbuser(namalengkap,iduser,password,level,status,user_ubah,tgl_ubah)values(nama,usern,pass,lev,stat,user,tgl);
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TampilBarang`()
BEGIN
START TRANSACTION;
SELECT tbjenis.nama as jenis,tbmerk.nama as merk,
tbbarang.* 
from tbbarang,tbjenis,tbmerk
 WHERE tbbarang.idjenis=tbjenis.idjenis
AND tbbarang.idmerk=tbmerk.idmerk
order by Tbbarang.kdbarang;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TampilKirim`()
BEGIN
START TRANSACTION;
SELECT tbtmpkirimun1.KdkirimUn,tbtmpkirimun1.tglkirim,tbtmpkirimun1.user_ubah,tbtmpkirimun1.tgl_ubah,
tbtmpkirimun1.kdunit,tbunit.namaunit 
FROM tbtmpkirimun1,TBTMPKIRIMUN,tbunit
WHERE tbtmpkirimun1.kdkirimUn=TBTMPKIRIMUN.kdkirimUn
AND tbUnit.kdUnit=tbtmpkirimun1.kdUnit 
and tbtmpkirimun1.flag='N'
GROUP BY tbtmpkirimun1.kdkirimUn;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TampilRetur`()
BEGIN
START TRANSACTION;
select tbtmpretursup1.*,
tbdistributor.namadistributor
from tbtmpretursup1,tbdistributor
where  tbtmpretursup1.kddistributor=tbdistributor.kddistributor
and tbtmpretursup1.flag='N'
group by tbtmpretursup1.kdretursup;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TampilReturU`()
BEGIN
START TRANSACTION;
select tbtmpreturun1.*,tbunit.namaunit
from tbtmpreturun1,tbUnit
where tbtmpreturun1.kdUnit=tbUnit.kdUnit
and tbtmpreturun1.flag='N'
group by tbtmpreturun1.kdreturUn;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TamTerima`()
BEGIN
START TRANSACTION;
SELECT tbterima.*,tbmerk.nama,tbdistributor.namaDistributor,tbdetterima.*,tbbarang.* 
FROM tbterima,tbdetterima,tbdistributor,tbbarang,tbmerk
WHERE tbterima.kdkirim=tbdetterima.kdkirim
AND tbdetterima.kdbarang=tbbarang.kdbarang
AND tbbarang.idmerk=tbmerk.idmerk
AND tbdistributor.kddistributor=tbterima.kddistributor 
GROUP BY tbterima.kdkirim;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TerimaDis`(IN `kode` VARCHAR(15))
BEGIN
START TRANSACTION;
SELECT tbtmpterima1.Kdkirim,tbtmpterima1.tgl_terima,tbtmpterima1.nofaktur,tbtmpterima1.tglfaktur,tbtmpterima1.user_ubah,tbtmpterima1.tgl_ubah,tbtmpterima1.kddistributor,tbdistributor.namaDistributor FROM tbtmpterima1,tbtmpterima,tbdistributor
WHERE tbtmpterima1.kdkirim=tbtmpterima.kdkirim
AND tbdistributor.kddistributor=tbtmpterima1.kddistributor
AND tbdistributor.kddistributor LIKE kode 
AND tbtmpterima1.flag='N'
GROUP BY tbtmpterima1.kdkirim;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TerimaKode`(IN `kode` VARCHAR(15))
BEGIN
START TRANSACTION;
SELECT tbtmpterima1.Kdkirim,tbtmpterima1.tgl_terima,tbtmpterima1.nofaktur,tbtmpterima1.tglfaktur,tbtmpterima1.user_ubah,tbtmpterima1.tgl_ubah,tbtmpterima1.kddistributor,tbdistributor.namaDistributor FROM tbtmpterima1,tbtmpterima,tbdistributor
WHERE tbtmpterima1.kdkirim=tbtmpterima.kdkirim
AND tbdistributor.kddistributor=tbtmpterima1.kddistributor
AND tbtmpterima1.kdkirim LIKE kode 
AND tbtmpterima1.flag='N'
GROUP BY tbtmpterima1.kdkirim;
COMMIT;
END$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `TerimaTerima`()
BEGIN
START TRANSACTION;
SELECT tbtmpterima1.Kdkirim,tbtmpterima1.tgl_terima,tbtmpterima1.nofaktur,tbtmpterima1.tglfaktur,tbtmpterima1.user_ubah,tbtmpterima1.tgl_ubah,tbtmpterima1.kddistributor,tbdistributor.namaDistributor FROM tbtmpterima1,tbtmpterima,tbdistributor
WHERE tbtmpterima1.kdkirim=tbtmpterima.kdkirim
AND tbdistributor.kddistributor=tbtmpterima1.kddistributor
AND tbtmpterima1.flag='N'
GROUP BY tbtmpterima1.kdkirim;
COMMIT;
END$$

DELIMITER ;

-- --------------------------------------------------------

--
-- Table structure for table `tbbarang`
--

CREATE TABLE IF NOT EXISTS `tbbarang` (
  `kdBarang` varchar(15) NOT NULL,
  `namaBarang` varchar(30) DEFAULT NULL,
  `satuan` varchar(10) DEFAULT NULL,
  `keterangan` text,
  `user_ubah` varchar(25) DEFAULT NULL,
  `tgl_ubah` datetime DEFAULT NULL,
  `idmerk` varchar(10) DEFAULT NULL,
  `idjenis` varchar(10) DEFAULT NULL,
  `Hargadasar` int(15) NOT NULL,
  `HargaFixed` int(15) NOT NULL,
  `stokAkhir` int(12) NOT NULL,
  PRIMARY KEY (`kdBarang`),
  KEY `idmerk` (`idmerk`),
  KEY `idjenis` (`idjenis`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbbarang`
--

INSERT INTO `tbbarang` (`kdBarang`, `namaBarang`, `satuan`, `keterangan`, `user_ubah`, `tgl_ubah`, `idmerk`, `idjenis`, `Hargadasar`, `HargaFixed`, `stokAkhir`) VALUES
('B-000001', 'SADS', 'PCS', '-', 'ADMIN', '2014-02-12 17:35:07', 'M-00001', 'J-00001', 10000, 12500, 7),
('B-000002', 'DASDAD', 'PCS', '-', 'ADMIN', '2014-02-12 17:35:21', 'M-00001', 'J-00007', 0, 0, 0);

-- --------------------------------------------------------

--
-- Table structure for table `tbdetkirimun`
--

CREATE TABLE IF NOT EXISTS `tbdetkirimun` (
  `kdkirimun` varchar(15) NOT NULL,
  `kdbarang` varchar(15) NOT NULL,
  `jml` int(12) NOT NULL,
  `harga` int(15) NOT NULL,
  `total` int(15) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbdetkirimun`
--

INSERT INTO `tbdetkirimun` (`kdkirimun`, `kdbarang`, `jml`, `harga`, `total`) VALUES
('0001/KU/02/14', 'B-000001', 5, 12500, 62500);

-- --------------------------------------------------------

--
-- Table structure for table `tbdetretursup`
--

CREATE TABLE IF NOT EXISTS `tbdetretursup` (
  `kdReturSup` varchar(15) NOT NULL,
  `kdbarang` varchar(15) NOT NULL,
  `harga` int(15) NOT NULL,
  `jml` int(12) NOT NULL,
  `total` int(15) NOT NULL,
  `alasan` text NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbdetretursup`
--

INSERT INTO `tbdetretursup` (`kdReturSup`, `kdbarang`, `harga`, `jml`, `total`, `alasan`) VALUES
('0001/RS/02/14', 'B-000001', 10000, 10, 100000, 'DDD');

-- --------------------------------------------------------

--
-- Table structure for table `tbdetreturun`
--

CREATE TABLE IF NOT EXISTS `tbdetreturun` (
  `kdReturUn` varchar(15) NOT NULL,
  `kdbarang` varchar(15) NOT NULL,
  `harga` int(15) NOT NULL,
  `jml` int(12) NOT NULL,
  `total` int(15) NOT NULL,
  `alasan` text NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `tbdetterima`
--

CREATE TABLE IF NOT EXISTS `tbdetterima` (
  `kdKirim` varchar(20) DEFAULT NULL,
  `kdBarang` varchar(15) DEFAULT NULL,
  `jumlah` int(12) DEFAULT NULL,
  `harga` int(15) DEFAULT NULL,
  `total` int(15) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbdetterima`
--

INSERT INTO `tbdetterima` (`kdKirim`, `kdBarang`, `jumlah`, `harga`, `total`) VALUES
('0001/BM/02/14', 'B-000001', 12, 10000, 120000),
('0002/BM/02/14', 'B-000001', 10, 10000, 100000);

-- --------------------------------------------------------

--
-- Table structure for table `tbdistributor`
--

CREATE TABLE IF NOT EXISTS `tbdistributor` (
  `kdDistributor` varchar(12) NOT NULL,
  `namaDistributor` varchar(30) NOT NULL,
  `alamat` text NOT NULL,
  `telp` int(13) NOT NULL,
  `kontakPerson` varchar(25) NOT NULL,
  `user_ubah` varchar(25) NOT NULL,
  `tgl_ubah` datetime NOT NULL,
  PRIMARY KEY (`kdDistributor`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbdistributor`
--

INSERT INTO `tbdistributor` (`kdDistributor`, `namaDistributor`, `alamat`, `telp`, `kontakPerson`, `user_ubah`, `tgl_ubah`) VALUES
('DS00001', 'DASDAS', 'SDA\r\n', 3442, 'AFADF', 'ADMIN', '2014-02-12 16:30:54'),
('DS00002', 'ADASDAS', 'SADAS\r\n', 342, 'SAFA', 'ADMIN', '2014-02-12 16:31:03'),
('DS00003', 'SADA', 'DSDA\r\n', 243, 'ASFFD', 'ADMIN', '2014-02-12 16:31:10'),
('DS00004', 'FSAF', 'FASF\r\n', 325, 'FAF', 'ADMIN', '2014-02-12 16:31:18'),
('DS00005', 'FASF', 'SFA\r\n', 323, 'FAF', 'ADMIN', '2014-02-12 16:31:23'),
('DS00006', 'FAF', 'SFA\r\n', 43243, 'SAFA', 'ADMIN', '2014-02-12 16:31:31'),
('DS00007', 'FSAFA', 'FASFA\r\n', 32, 'FSD', 'ADMIN', '2014-02-12 16:31:36'),
('DS00008', 'FSF', 'FSD\r\n', 534, 'FSFD', 'ADMIN', '2014-02-12 16:31:43'),
('DS00009', 'FA', 'SFA\r\n', 325, 'FSDFS', 'ADMIN', '2014-02-12 16:31:49'),
('DS00010', 'RFF', 'DSFSD\r\n', 534, 'GDG', 'ADMIN', '2014-02-12 16:31:57'),
('DS00011', 'FSDFD', 'FSFD\r\n', 5345, 'FSDFD', 'ADMIN', '2014-02-12 16:32:05');

-- --------------------------------------------------------

--
-- Table structure for table `tbjenis`
--

CREATE TABLE IF NOT EXISTS `tbjenis` (
  `idjenis` varchar(10) NOT NULL,
  `nama` varchar(25) NOT NULL,
  `user_ubah` varchar(20) NOT NULL,
  `tgl_ubah` datetime NOT NULL,
  PRIMARY KEY (`idjenis`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbjenis`
--

INSERT INTO `tbjenis` (`idjenis`, `nama`, `user_ubah`, `tgl_ubah`) VALUES
('J-00001', 'MSDAD', 'ADMIN', '2014-02-12 16:15:44'),
('J-00002', 'SDASDS', 'ADMIN', '2014-02-12 16:15:49'),
('J-00003', 'DAADA', 'ADMIN', '2014-02-12 16:15:54'),
('J-00004', 'DAD', 'ADMIN', '2014-02-12 16:15:59'),
('J-00005', 'DSA', 'ADMIN', '2014-02-12 16:16:03'),
('J-00006', 'AAA', 'ADMIN', '2014-02-12 16:16:06'),
('J-00007', 'AA', 'ADMIN', '2014-02-12 16:16:11'),
('J-00008', 'SADADAD', 'ADMIN', '2014-02-12 16:16:16'),
('J-00009', 'ADADADADD', 'ADMIN', '2014-02-12 16:16:21'),
('J-00010', 'FDFS', 'ADMIN', '2014-02-12 16:23:04'),
('J-00011', 'DFSFSF', 'ADMIN', '2014-02-12 16:23:11');

-- --------------------------------------------------------

--
-- Table structure for table `tbkirimun`
--

CREATE TABLE IF NOT EXISTS `tbkirimun` (
  `kdKirimUn` varchar(15) NOT NULL,
  `tglkirim` date DEFAULT NULL,
  `kdunit` varchar(15) DEFAULT NULL,
  `user_ubah` varchar(25) DEFAULT NULL,
  `tgl_ubah` datetime DEFAULT NULL,
  `konfirm` varchar(5) DEFAULT NULL,
  PRIMARY KEY (`kdKirimUn`),
  KEY `kdunit` (`kdunit`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbkirimun`
--

INSERT INTO `tbkirimun` (`kdKirimUn`, `tglkirim`, `kdunit`, `user_ubah`, `tgl_ubah`, `konfirm`) VALUES
('0001/KU/02/14', '2014-02-12', 'UN-01', 'ADMIN', '2014-02-12 00:00:00', 'Y');

-- --------------------------------------------------------

--
-- Table structure for table `tbmerk`
--

CREATE TABLE IF NOT EXISTS `tbmerk` (
  `idmerk` varchar(10) NOT NULL,
  `nama` varchar(25) NOT NULL,
  `user_ubah` varchar(20) NOT NULL,
  `tgl_ubah` datetime NOT NULL,
  PRIMARY KEY (`idmerk`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbmerk`
--

INSERT INTO `tbmerk` (`idmerk`, `nama`, `user_ubah`, `tgl_ubah`) VALUES
('M-00001', 'SADSA', 'ADMIN', '2014-02-12 16:25:12'),
('M-00002', 'DASDAD', 'ADMIN', '2014-02-12 16:25:16'),
('M-00003', 'DASDAA', 'ADMIN', '2014-02-12 16:25:20'),
('M-00004', 'D', 'ADMIN', '2014-02-12 16:25:24'),
('M-00005', 'DD', 'ADMIN', '2014-02-12 16:25:27'),
('M-00006', 'DDDD', 'ADMIN', '2014-02-12 16:25:31'),
('M-00007', 'DDDDD', 'ADMIN', '2014-02-12 16:25:35'),
('M-00008', 'DDDDDD', 'ADMIN', '2014-02-12 16:25:42'),
('M-00009', 'DDDDDDD', 'ADMIN', '2014-02-12 16:25:47'),
('M-00010', 'DSDASDADSFD', 'ADMIN', '2014-02-12 16:25:52'),
('M-00011', 'SDADADADASD', 'ADMIN', '2014-02-12 16:25:57'),
('M-00012', 'DSDADA', 'ADMIN', '2014-02-12 16:26:00');

-- --------------------------------------------------------

--
-- Table structure for table `tbretursup`
--

CREATE TABLE IF NOT EXISTS `tbretursup` (
  `kdReturSup` varchar(15) NOT NULL,
  `tglretur` date NOT NULL,
  `kddistributor` varchar(12) NOT NULL,
  `user_ubah` varchar(25) NOT NULL,
  `tgl_ubah` datetime NOT NULL,
  `konfirm` varchar(1) NOT NULL,
  PRIMARY KEY (`kdReturSup`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbretursup`
--

INSERT INTO `tbretursup` (`kdReturSup`, `tglretur`, `kddistributor`, `user_ubah`, `tgl_ubah`, `konfirm`) VALUES
('0001/RS/02/14', '2014-02-12', 'DS00004', 'ADMIN', '2014-02-12 18:55:41', 'Y');

-- --------------------------------------------------------

--
-- Table structure for table `tbreturun`
--

CREATE TABLE IF NOT EXISTS `tbreturun` (
  `kdReturUn` varchar(15) NOT NULL,
  `tglreturUn` date NOT NULL,
  `kdunit` varchar(12) NOT NULL,
  `user_ubah` varchar(25) NOT NULL,
  `tgl_ubah` datetime NOT NULL,
  `konfirm` varchar(1) NOT NULL,
  PRIMARY KEY (`kdReturUn`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `tbstok`
--

CREATE TABLE IF NOT EXISTS `tbstok` (
  `idstok` int(10) NOT NULL AUTO_INCREMENT,
  `kdbarang` varchar(15) DEFAULT NULL,
  `no_bukti` varchar(15) NOT NULL,
  `masuk` int(8) DEFAULT NULL,
  `keluar` int(8) DEFAULT NULL,
  `stok` int(8) DEFAULT NULL,
  `keterangan` varchar(20) NOT NULL,
  `user_ubah` varchar(25) DEFAULT NULL,
  `tgl_ubah` datetime DEFAULT NULL,
  PRIMARY KEY (`idstok`),
  KEY `kdbarang` (`kdbarang`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=8 ;

--
-- Dumping data for table `tbstok`
--

INSERT INTO `tbstok` (`idstok`, `kdbarang`, `no_bukti`, `masuk`, `keluar`, `stok`, `keterangan`, `user_ubah`, `tgl_ubah`) VALUES
(2, 'B-000001', '-', 0, 0, 0, 'INPUT BARANG', 'ADMIN', '2014-02-12 17:35:07'),
(3, 'B-000002', '-', 0, 0, 0, 'INPUT BARANG', 'ADMIN', '2014-02-12 17:35:21'),
(4, 'B-000001', '0001/BM/02/14', 12, 0, 12, 'TERIMA DISTRIBUTOR', 'ADMIN', '2014-02-12 18:51:06'),
(5, 'B-000001', '0002/BM/02/14', 10, 0, 22, 'TERIMA DISTRIBUTOR', 'ADMIN', '2014-02-12 18:51:40'),
(6, 'B-000001', '0001/KU/02/14', 0, 5, 17, 'KIRIM UNIT', 'ADMIN', '2014-02-12 18:52:52'),
(7, 'B-000001', '0001/RS/02/14', 0, 100000, 7, 'RETUR DISTRIBUTOR', 'ADMIN', '2014-02-12 18:55:41');

-- --------------------------------------------------------

--
-- Table structure for table `tbterima`
--

CREATE TABLE IF NOT EXISTS `tbterima` (
  `kdKirim` varchar(20) NOT NULL,
  `tgl_terima` date DEFAULT NULL,
  `nofaktur` varchar(20) DEFAULT NULL,
  `tglfaktur` date DEFAULT NULL,
  `kdDistributor` varchar(15) DEFAULT NULL,
  `user_ubah` varchar(25) DEFAULT NULL,
  `tgl_ubah` datetime DEFAULT NULL,
  `konfirm` varchar(1) NOT NULL,
  PRIMARY KEY (`kdKirim`),
  KEY `kdDistributor` (`kdDistributor`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbterima`
--

INSERT INTO `tbterima` (`kdKirim`, `tgl_terima`, `nofaktur`, `tglfaktur`, `kdDistributor`, `user_ubah`, `tgl_ubah`, `konfirm`) VALUES
('0001/BM/02/14', '2014-02-12', 'F001', '2014-02-12', 'DS00004', 'ADMIN', '2014-02-12 18:51:06', 'Y'),
('0002/BM/02/14', '2014-02-12', 'F9', '2014-02-12', 'DS00001', 'ADMIN', '2014-02-12 18:51:40', 'Y');

-- --------------------------------------------------------

--
-- Table structure for table `tbtmpkirimun`
--

CREATE TABLE IF NOT EXISTS `tbtmpkirimun` (
  `kdkirimun` varchar(15) NOT NULL,
  `kdbarang` varchar(15) NOT NULL,
  `jml` int(12) NOT NULL,
  `harga` int(15) NOT NULL,
  `total` int(15) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbtmpkirimun`
--

INSERT INTO `tbtmpkirimun` (`kdkirimun`, `kdbarang`, `jml`, `harga`, `total`) VALUES
('0001/KU/02/14', 'B-000001', 5, 12500, 62500),
('0002/KU/02/14', 'B-000001', 2, 12500, 25000);

-- --------------------------------------------------------

--
-- Table structure for table `tbtmpkirimun1`
--

CREATE TABLE IF NOT EXISTS `tbtmpkirimun1` (
  `kdKirimUn` varchar(15) NOT NULL,
  `tglkirim` date DEFAULT NULL,
  `kdunit` varchar(15) DEFAULT NULL,
  `user_ubah` varchar(25) DEFAULT NULL,
  `tgl_ubah` datetime DEFAULT NULL,
  `flag` varchar(1) NOT NULL DEFAULT 'N',
  PRIMARY KEY (`kdKirimUn`),
  KEY `kdunit` (`kdunit`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbtmpkirimun1`
--

INSERT INTO `tbtmpkirimun1` (`kdKirimUn`, `tglkirim`, `kdunit`, `user_ubah`, `tgl_ubah`, `flag`) VALUES
('0001/KU/02/14', '2014-02-12', 'UN-01', 'ADMIN', '2014-02-12 00:00:00', 'Y'),
('0002/KU/02/14', '2014-02-13', 'UN-01', 'ADMIN', '2014-02-13 18:53:06', 'N');

-- --------------------------------------------------------

--
-- Table structure for table `tbtmpretursup`
--

CREATE TABLE IF NOT EXISTS `tbtmpretursup` (
  `kdReturSup` varchar(15) NOT NULL,
  `kdbarang` varchar(15) NOT NULL,
  `harga` int(15) NOT NULL,
  `jml` int(12) NOT NULL,
  `total` int(15) NOT NULL,
  `alasan` text NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbtmpretursup`
--

INSERT INTO `tbtmpretursup` (`kdReturSup`, `kdbarang`, `harga`, `jml`, `total`, `alasan`) VALUES
('0001/RS/02/14', 'B-000001', 10000, 10, 100000, 'DDD'),
('0002/RS/02/14', 'B-000001', 10000, 2, 20000, 'DDDD');

-- --------------------------------------------------------

--
-- Table structure for table `tbtmpretursup1`
--

CREATE TABLE IF NOT EXISTS `tbtmpretursup1` (
  `kdReturSup` varchar(15) NOT NULL,
  `tglretur` date NOT NULL,
  `kddistributor` varchar(12) NOT NULL,
  `user_ubah` varchar(25) NOT NULL,
  `tgl_ubah` datetime NOT NULL,
  `flag` varchar(1) NOT NULL DEFAULT 'N',
  PRIMARY KEY (`kdReturSup`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbtmpretursup1`
--

INSERT INTO `tbtmpretursup1` (`kdReturSup`, `tglretur`, `kddistributor`, `user_ubah`, `tgl_ubah`, `flag`) VALUES
('0001/RS/02/14', '2014-02-12', 'DS00004', 'ADMIN', '2014-02-12 18:55:32', 'Y'),
('0002/RS/02/14', '2014-02-13', 'DS00004', 'ADMIN', '2014-02-13 19:08:16', 'N');

-- --------------------------------------------------------

--
-- Table structure for table `tbtmpreturun`
--

CREATE TABLE IF NOT EXISTS `tbtmpreturun` (
  `kdReturUn` varchar(15) NOT NULL,
  `kdbarang` varchar(15) NOT NULL,
  `harga` int(15) NOT NULL,
  `jml` int(12) NOT NULL,
  `total` int(15) NOT NULL,
  `alasan` text NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbtmpreturun`
--

INSERT INTO `tbtmpreturun` (`kdReturUn`, `kdbarang`, `harga`, `jml`, `total`, `alasan`) VALUES
('0001/RU/02/14', 'B-000001', 10000, 2, 20000, 'SADASDAD');

-- --------------------------------------------------------

--
-- Table structure for table `tbtmpreturun1`
--

CREATE TABLE IF NOT EXISTS `tbtmpreturun1` (
  `kdReturUn` varchar(15) NOT NULL,
  `tglreturUn` date NOT NULL,
  `kdunit` varchar(12) NOT NULL,
  `user_ubah` varchar(25) NOT NULL,
  `tgl_ubah` datetime NOT NULL,
  `FLAG` varchar(1) NOT NULL DEFAULT 'N',
  PRIMARY KEY (`kdReturUn`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbtmpreturun1`
--

INSERT INTO `tbtmpreturun1` (`kdReturUn`, `tglreturUn`, `kdunit`, `user_ubah`, `tgl_ubah`, `FLAG`) VALUES
('0001/RU/02/14', '2014-02-13', 'UN-01', 'ADMIN', '2014-02-13 19:06:15', 'N');

-- --------------------------------------------------------

--
-- Table structure for table `tbtmpterima`
--

CREATE TABLE IF NOT EXISTS `tbtmpterima` (
  `kdKirim` varchar(15) DEFAULT NULL,
  `kdBarang` varchar(15) DEFAULT NULL,
  `jumlah` int(12) DEFAULT NULL,
  `hargaDasar` int(15) DEFAULT NULL,
  `totalDasar` int(15) DEFAULT NULL,
  `HargaFixed` int(15) NOT NULL,
  `persen` int(3) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbtmpterima`
--

INSERT INTO `tbtmpterima` (`kdKirim`, `kdBarang`, `jumlah`, `hargaDasar`, `totalDasar`, `HargaFixed`, `persen`) VALUES
('0001/BM/02/14', 'B-000001', 12, 10000, 120000, 12500, 25),
('0002/BM/02/14', 'B-000001', 10, 10000, 100000, 12500, 25),
('0003/BM/02/14', 'B-000001', 2, 10000, 20000, 12500, 25);

-- --------------------------------------------------------

--
-- Table structure for table `tbtmpterima1`
--

CREATE TABLE IF NOT EXISTS `tbtmpterima1` (
  `kdKirim` varchar(20) NOT NULL,
  `tgl_terima` date DEFAULT NULL,
  `nofaktur` varchar(20) DEFAULT NULL,
  `tglfaktur` varchar(14) DEFAULT NULL,
  `kdDistributor` varchar(15) DEFAULT NULL,
  `user_ubah` varchar(25) DEFAULT NULL,
  `tgl_ubah` datetime DEFAULT NULL,
  `flag` varchar(1) NOT NULL DEFAULT 'N',
  PRIMARY KEY (`kdKirim`),
  KEY `kdDistributor` (`kdDistributor`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbtmpterima1`
--

INSERT INTO `tbtmpterima1` (`kdKirim`, `tgl_terima`, `nofaktur`, `tglfaktur`, `kdDistributor`, `user_ubah`, `tgl_ubah`, `flag`) VALUES
('0001/BM/02/14', '2014-02-12', 'F001', '2014-02-12', 'DS00004', 'ADMIN', '2014-02-12 17:50:09', 'Y'),
('0002/BM/02/14', '2014-02-12', 'F9', '2014-02-12', 'DS00001', 'ADMIN', '2014-02-12 18:10:01', 'Y'),
('0003/BM/02/14', '2014-02-13', 'F001', '2014-02-13', 'DS00004', 'ADMIN', '2014-02-13 18:19:38', 'N');

-- --------------------------------------------------------

--
-- Table structure for table `tbunit`
--

CREATE TABLE IF NOT EXISTS `tbunit` (
  `KdUnit` varchar(15) NOT NULL,
  `namaUnit` varchar(25) DEFAULT NULL,
  `alamat` text,
  `telp` int(13) DEFAULT NULL,
  `kontakPerson` varchar(25) DEFAULT NULL,
  `user_ubah` varchar(25) DEFAULT NULL,
  `tgl_ubah` datetime DEFAULT NULL,
  PRIMARY KEY (`KdUnit`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbunit`
--

INSERT INTO `tbunit` (`KdUnit`, `namaUnit`, `alamat`, `telp`, `kontakPerson`, `user_ubah`, `tgl_ubah`) VALUES
('UN-01', 'MERIAM', 'JL. CIBATU\r\n', 2147483647, 'MANDI', 'ADMIN', '2014-02-12 15:51:00');

-- --------------------------------------------------------

--
-- Table structure for table `tbuser`
--

CREATE TABLE IF NOT EXISTS `tbuser` (
  `iduser` varchar(20) NOT NULL,
  `namalengkap` varchar(30) NOT NULL,
  `password` varchar(100) NOT NULL,
  `user_ubah` varchar(20) NOT NULL,
  `tgl_ubah` datetime NOT NULL,
  `level` varchar(15) NOT NULL,
  `status` varchar(1) NOT NULL,
  PRIMARY KEY (`iduser`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbuser`
--

INSERT INTO `tbuser` (`iduser`, `namalengkap`, `password`, `user_ubah`, `tgl_ubah`, `level`, `status`) VALUES
('ADMIN', 'ADMINISTRATOR', 'c4ca4238a0b923820dcc509a6f75849b', 'admin', '2013-12-23 00:00:00', 'ADMINISTRATOR', 'Y'),
('AWAL', 'AWAL', '202cb962ac59075b964b07152d234b70', 'ADMIN', '2014-01-25 20:37:59', 'ADMINISTRATOR', 'N');

--
-- Constraints for dumped tables
--

--
-- Constraints for table `tbbarang`
--
ALTER TABLE `tbbarang`
  ADD CONSTRAINT `tbbarang_ibfk_1` FOREIGN KEY (`idmerk`) REFERENCES `tbmerk` (`idmerk`) ON UPDATE CASCADE,
  ADD CONSTRAINT `tbbarang_ibfk_2` FOREIGN KEY (`idjenis`) REFERENCES `tbjenis` (`idjenis`) ON UPDATE CASCADE;

--
-- Constraints for table `tbkirimun`
--
ALTER TABLE `tbkirimun`
  ADD CONSTRAINT `tbkirimun_ibfk_1` FOREIGN KEY (`kdunit`) REFERENCES `tbunit` (`KdUnit`) ON UPDATE CASCADE;

--
-- Constraints for table `tbstok`
--
ALTER TABLE `tbstok`
  ADD CONSTRAINT `tbstok_ibfk_1` FOREIGN KEY (`kdbarang`) REFERENCES `tbbarang` (`kdBarang`) ON UPDATE CASCADE;

--
-- Constraints for table `tbterima`
--
ALTER TABLE `tbterima`
  ADD CONSTRAINT `tbterima_ibfk_1` FOREIGN KEY (`kdDistributor`) REFERENCES `tbdistributor` (`kdDistributor`) ON UPDATE CASCADE;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;

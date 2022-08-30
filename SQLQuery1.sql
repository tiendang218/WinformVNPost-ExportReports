drop database ChungTuHCC
create database ChungTuHCC
 use [ChungTuHCC]
SELECT FORMAT (getdate(), 'dd/MM/yyyy ') as date
CREATE TABLE CT_HCC (
    SO_CT	int PRIMARY KEY not null,
	Approved bit not null,
TEN_NGUOI_NHAN	nvarchar(50) not null ,	
DIA_CHI	nvarchar(50) not null,	
SO_HS_HCC	nvarchar(50) not null,	
NGAY_HEN	date	not null,
NGAY_NHAN	date	not null,
DIA_CHI_THUONG_TRU nvarchar(50) not null,
TINH_NHAN	int	not null,
HUYEN_NHAN	int	not null,
XA_NHAN	int	not null,
SO_HS_KEM	varchar(10)	not null,
DIEN_THOAI	varchar(50)	not null,
MA_LOAI_HS	int	not null,
TRONG_LUONG	varchar(10)not null,
MA_BUUGUI	varchar(10)not null	,
GHI_CHU	nvarchar(50)not null,
NGAY_DUYET date NULL,
CUOC_PHI decimal(15) NOT NULL
);
--select NGAY_NHAN from CT_HCC
--Go
----dd/mm/yyyy format
--select convert (varchar,NGAY_NHAN,103) from CT_HCC
--select convert (varchar,NGAY_DUYET,103) from CT_HCC
--select convert (varchar,NGAY_HEN,103) from CT_HCC
--Go
ALTER TABLE CT_HCC
ALTER COLUMN NGAY_HEN nvarchar(10) not null;
ALTER TABLE CT_HCC
ALTER COLUMN NGAY_NHAN nvarchar(10) not null;
ALTER TABLE CT_HCC
ALTER COLUMN NGAY_DUYET nvarchar(10) not null;

CREATE TABLE LOAI_HS
(
MA_LOAI_HS int PRIMARY KEY NOT NULL,
TEN_LOAI_HS	nvarchar(50) NOT NULL,
);
CREATE TABLE TINH
(
MA_TINH int PRIMARY KEY NOT NULL,
TEN_TINH	nvarchar(50) NOT NULL,
);
CREATE TABLE HUYEN_TP
(
MA_HUYENTP int PRIMARY KEY NOT NULL,
TEN_HUYENTP	nvarchar(50) NOT NULL,
MA_TINH int NOT NULL,
);
CREATE TABLE XAPHUONG
(
MA_XA_PHUONG int PRIMARY KEY NOT NULL,
TEN_XA_PHUONG	nvarchar(50) NOT NULL,
MA_HUYENTP int NOT NULL,
);
ALTER TABLE CT_HCC
ADD CONSTRAINT FK_tinh
FOREIGN KEY (TINH_NHAN) REFERENCES TINH(MA_TINH);

ALTER TABLE CT_HCC
ADD CONSTRAINT FK_HUYEN
FOREIGN KEY (HUYEN_NHAN) REFERENCES HUYEN_TP(MA_HUYENTP);

ALTER TABLE CT_HCC
ADD CONSTRAINT FK_XA
FOREIGN KEY (XA_NHAN) REFERENCES XAPHUONG(MA_XA_PHUONG);

ALTER TABLE CT_HCC
ADD CONSTRAINT FK_CT_HCC
FOREIGN KEY (MA_LOAI_HS) REFERENCES LOAI_HS(MA_LOAI_HS);


DROP PROCEDURE [CRUD]

USE [ChungTuHCC]  
GO  
SET ANSI_NULLS ON  
GO  
SET QUOTED_IDENTIFIER ON  
GO  

create PROCEDURE [CRUD]  
       @SO_CT	int,
	   @Approved bit,
@TEN_NGUOI_NHAN	nvarchar(50) ,	
@DIA_CHI	nvarchar(50),	
@SO_HS_HCC	nvarchar(50),	
@NGAY_HEN	nvarchar(10)	,
@NGAY_NHAN	nvarchar(10)	,
@DIA_CHI_THUONG_TRU nvarchar(50),
@TINH_NHAN	int,
@HUYEN_NHAN	int,
@XA_NHAN	int,
@SO_HS_KEM	varchar(10),
@DIEN_THOAI	varchar(50),
@MA_LOAI_HS	int,
@TRONG_LUONG	varchar(10),
@MA_BUUGUI	varchar(10),
@GHI_CHU	nvarchar(50),
@NGAY_DUYET nvarchar(10),
@CUOC_PHI decimal(10,2),

@OperationType int   
    --================================================  
    --operation types   
    -- 1) Insert  
    -- 2) Update  
    -- 3) Delete  
    -- 4) Select Perticular Record  
    -- 5) Selec All  
AS  
begin
    -- SET NOCOUNT ON added to prevent extra result sets from  
    -- interfering with SELECT statements.  
    SET NOCOUNT ON;  
      
    --select operation  
    IF @OperationType=1  
    BEGIN  
        INSERT INTO CT_HCC VALUES (
		@SO_CT,
		@Approved,
		@TEN_NGUOI_NHAN,
		@DIA_CHI,
		@SO_HS_HCC,
		@NGAY_HEN,
		@NGAY_NHAN,
		@DIA_CHI_THUONG_TRU,
		@TINH_NHAN,
		@HUYEN_NHAN,
		@XA_NHAN,
		@SO_HS_KEM,
		@DIEN_THOAI,
		@MA_LOAI_HS,
		@TRONG_LUONG,
		@MA_BUUGUI,
		@GHI_CHU,
		@NGAY_DUYET,
		@CUOC_PHI)  
    END  
    ELSE IF @OperationType=2  
    BEGIN  
        UPDATE CT_HCC SET 
		SO_CT=@SO_CT,
		Approved=@Approved,
		TEN_NGUOI_NHAN=@TEN_NGUOI_NHAN,
		DIA_CHI=@DIA_CHI,
		SO_HS_HCC=@SO_HS_HCC,
		NGAY_HEN=@NGAY_HEN,
		NGAY_NHAN=@NGAY_NHAN,
		DIA_CHI_THUONG_TRU=@DIA_CHI_THUONG_TRU,
		TINH_NHAN=@TINH_NHAN,
		HUYEN_NHAN=@HUYEN_NHAN,
		XA_NHAN=@XA_NHAN,
		SO_HS_KEM=@SO_HS_KEM,
		DIEN_THOAI=@DIEN_THOAI,
		MA_LOAI_HS=@MA_LOAI_HS,
		TRONG_LUONG=@TRONG_LUONG,
		MA_BUUGUI=@MA_BUUGUI,
		GHI_CHU=@GHI_CHU,
		NGAY_DUYET=@NGAY_DUYET,
		CUOC_PHI=@CUOC_PHI
		WHERE SO_CT=@SO_CT  
    END  
    ELSE IF @OperationType=3  
    BEGIN  
        DELETE FROM CT_HCC WHERE SO_CT=@SO_CT  
    END  
    ELSE IF @OperationType=4  
    BEGIN  
        SELECT * FROM CT_HCC WHERE SO_CT=@SO_CT  
    END  
	ELSE IF @OperationType=5 
    BEGIN  
        SELECT TEN_TINH,MA_TINH FROM TINH
    END
	ELSE IF @OperationType=6
	BEGIN  
		
        SELECT TEN_HUYENTP FROM HUYEN_TP as H, TINH as T where H.MA_TINH= T.MA_TINH
    END  
	ELSE IF @OperationType=7
	BEGIN  
        SELECT TEN_XA_PHUONG FROM XAPHUONG
    END  
	ELSE IF @OperationType=8
	BEGIN  
        SELECT TEN_LOAI_HS,MA_LOAI_HS FROM LOAI_HS
    END  
    ELSE   
    BEGIN  
        SELECT * FROM CT_HCC   
    END  
END  
select * from XAPHUONG
drop PROCEDURE dem_ban_ghi
CREATE PROCEDURE dem_ban_ghi @tinh int, @huyen int
AS
SELECT count(*) as "so_ban_ghi" FROM CT_HCC WHERE TINH_NHAN = @tinh and HUYEN_NHAN=@huyen and Approved= 1

EXEC dem_ban_ghi @tinh = 86, @huyen = 1

select * from XAPHUONG where MA_HUYENTP = 1

drop procedure auto_add_id
create procedure auto_add_id
as
DECLARE @dem INT
SELECT @dem = COUNT(*)
FROM CT_HCC
IF @dem = 0
BEGIN
	SELECT 1 as Số_CT
END
ELSE
BEGIN
	select MAX(SO_CT)+1 as Số_CT from CT_HCC
END

drop procedure bao_cao_cthcc
create procedure bao_cao_cthcc @ngayduyet varchar(10)
as 
with CountData as
(
select t.TEN_TINH as tinh ,H.TEN_HUYENTP as huyen, count(ct.SO_CT) as So_bao_cao
from CT_HCC as ct, HUYEN_TP as H, TINH as t where ct.HUYEN_NHAN= H.MA_HUYENTP and h.MA_TINH=t.MA_TINH and ct.Approved=1 and ct.NGAY_DUYET=@ngayduyet
group by t.TEN_TINH, H.TEN_HUYENTP
--,ct.HUYEN_NHAN
)
 select tinh, huyen, So_bao_cao from CountData
 UNION ALL
 select '',N'Tổng Cộng', SUM(So_bao_cao)
 from CountData

exec bao_cao_cthcc @ngayduyet = '2022-08-22'

drop PROCEDURE load_huyen
CREATE PROCEDURE load_huyen @tinh int
AS
SELECT MA_HUYENTP,TEN_HUYENTP FROM HUYEN_TP WHERE MA_TINH=@tinh

EXEC load_huyen @tinh = 86

drop PROCEDURE load_xa
CREATE PROCEDURE load_xa @huyen int
AS
SELECT MA_XA_PHUONG,TEN_XA_PHUONG FROM XAPHUONG WHERE MA_HUYENTP=@huyen
EXEC load_xa @huyen = 2

drop procedure load_tinh_huyen_xa
create procedure load_tinh_huyen_xa @tinh int,@huyen int,@xa int
as
select t.TEN_TINH, h.TEN_HUYENTP, x.TEN_XA_PHUONG
from TINH as t, HUYEN_TP as h, XAPHUONG as x
where t.MA_TINH=@tinh and t.MA_TINH=h.MA_TINH and h.MA_HUYENTP= @huyen and x.MA_HUYENTP= h.MA_HUYENTP and x.MA_XA_PHUONG=@xa

drop proc load_loai_hoso @id=1
create proc load_loai_hoso @id int
as
select TEN_LOAI_HS from LOAI_HS
where MA_LOAI_HS=@id

drop proc duyet_ct
create proc duyet_ct @i int,@id int,@date varchar(10)
as
update CT_HCC set Approved=@i,NGAY_DUYET= @date
where SO_CT=@id


insert into LOAI_HS values(0,N'')
insert into LOAI_HS values(1,N'Cấp mới')
insert into LOAI_HS values(2,N'Mất cấp lại')
insert into LOAI_HS values(3,N'Đổi điều chỉnh')
insert into LOAI_HS values(4,N'Không chọn')

insert into TINH values(0,N'')
insert into TINH values(86,N'Bình Thuận')
insert into HUYEN_TP values(0,N'',0)
insert into HUYEN_TP values(1,N'TP.Phan Thiết',86)
insert into HUYEN_TP values(2,N'Tuy Phong',86)
insert into HUYEN_TP values(3,N'Hàm Thuận Bắc',86)
insert into HUYEN_TP values(4,N'Hàm Thuận Nam',86)
insert into HUYEN_TP values(5,N'Hàm Tân',86)
insert into HUYEN_TP values(6,N'Tánh Linh',86)
insert into HUYEN_TP values(7,N'Đức Linh',86)
insert into HUYEN_TP values(8,N'Bắc Bình',86)
insert into HUYEN_TP values(9,N'Lagi',86)
insert into HUYEN_TP values(10,N'Phú Quý',86)

insert into XAPHUONG values(0,N'',0)
insert into XAPHUONG values(1,N'Mũi Né',1)
insert into XAPHUONG values(2,N'Hàm Tiến',1)
insert into XAPHUONG values(3,N'Phú Hài',1)
insert into XAPHUONG values(4,N'Phú Thủy',1)
insert into XAPHUONG values(5,N'Phú Tài',1)
insert into XAPHUONG values(6,N'Phú Trinh',1)
insert into XAPHUONG values(7,N'Xuân An',1)
insert into XAPHUONG values(8,N'Thanh Hải',1)
insert into XAPHUONG values(9,N'Bình Hưng',1)
insert into XAPHUONG values(10,N'Đức Nghĩa',1)
insert into XAPHUONG values(11,N'Lạc Đạo',1)
insert into XAPHUONG values(12,N'Đức Thắng',1)
insert into XAPHUONG values(13,N'Hưng Long',1)
insert into XAPHUONG values(14,N'Đức Long',1)
insert into XAPHUONG values(15,N'Thiện Nghiệp',1)
insert into XAPHUONG values(16,N'Phong Nẫm',1)
insert into XAPHUONG values(17,N'Tiến Lợi',1)
insert into XAPHUONG values(18,N'Tiến Thành',1)

insert into XAPHUONG values(19,N'Liên Hương',2)
insert into XAPHUONG values(20,N'Phan Rí Cửa',2)
insert into XAPHUONG values(21,N'Phan Dũng',2)
insert into XAPHUONG values(22,N'Phong Phú',2)
insert into XAPHUONG values(23,N'Vĩnh Hảo',2)
insert into XAPHUONG values(24,N'Vĩnh Tân',2)
insert into XAPHUONG values(25,N'Phú Lạc',2)
insert into XAPHUONG values(26,N'Phước Thể',2)
insert into XAPHUONG values(27,N'Hòa Minh',2)
insert into XAPHUONG values(28,N'Chí Công',2)
insert into XAPHUONG values(29,N'Bình Thạnh',2)

insert into XAPHUONG values(30,N'Ma Lâm',3)
insert into XAPHUONG values(31,N'Phú Long',3)
insert into XAPHUONG values(32,N'La Dạ',3)
insert into XAPHUONG values(33,N'Đông Tiến',3)
insert into XAPHUONG values(34,N'Thuận Hòa',3)
insert into XAPHUONG values(35,N'Đông Giang',3)
insert into XAPHUONG values(36,N'Hàm Phú',3)
insert into XAPHUONG values(37,N'Hồng Liêm',3)
insert into XAPHUONG values(38,N'Thuận Minh',3)
insert into XAPHUONG values(39,N'Hồng Sơn',3)
insert into XAPHUONG values(40,N'Hàm Trí',3)
insert into XAPHUONG values(41,N'Hàm Đức',3)
insert into XAPHUONG values(42,N'Hàm Liêm',3)
insert into XAPHUONG values(43,N'Hàm Chính',3)
insert into XAPHUONG values(44,N'Hàm Hiệp',3)
insert into XAPHUONG values(45,N'Hàm Thắng',3)
insert into XAPHUONG values(46,N'Đa Mi',3)

insert into XAPHUONG values(47,N'Thuận Nam',4)
insert into XAPHUONG values(48,N'Mỹ Thạnh',4)
insert into XAPHUONG values(49,N'Hàm Cần',4)
insert into XAPHUONG values(50,N'Mương Mán',4)
insert into XAPHUONG values(51,N'Hàm Thạnh',4)
insert into XAPHUONG values(52,N'Hàm Kiệm',4)
insert into XAPHUONG values(53,N'Hàm Cường',4)
insert into XAPHUONG values(54,N'Hàm Mỹ',4)
insert into XAPHUONG values(55,N'Tân Lập',4)
insert into XAPHUONG values(56,N'Hàm Minh',4)
insert into XAPHUONG values(57,N'Thuận Quý',4)
insert into XAPHUONG values(58,N'Tân Thuận',4)
insert into XAPHUONG values(59,N'Tân Thành',4)

insert into XAPHUONG values(60,N'Tân Minh',5)
insert into XAPHUONG values(61,N'Tân Nghĩa',5)
insert into XAPHUONG values(62,N'Sông Phan',5)
insert into XAPHUONG values(63,N'Tân Phúc',5)
insert into XAPHUONG values(64,N'Tân Đức',5)
insert into XAPHUONG values(65,N'Tân Thắng',5)
insert into XAPHUONG values(66,N'Thắng Hải',5)
insert into XAPHUONG values(67,N'Tân Hà',5)
insert into XAPHUONG values(68,N'Tân Xuân',5)
insert into XAPHUONG values(69,N'Sơn Mỹ',5)

insert into XAPHUONG values(70,N'Lạc Tánh',6)
insert into XAPHUONG values(71,N'Bắc Ruộng',6)
insert into XAPHUONG values(72,N'Nghị Đức',6)
insert into XAPHUONG values(73,N'La Ngâu',6)
insert into XAPHUONG values(74,N'Huy Khiêm',6)
insert into XAPHUONG values(75,N'Măng Tố',6)
insert into XAPHUONG values(76,N'Đức Phú',6)
insert into XAPHUONG values(77,N'Đồng Kho',6)
insert into XAPHUONG values(78,N'Gia An',6)
insert into XAPHUONG values(79,N'Đức Bình',6)
insert into XAPHUONG values(80,N'Gia Huynh',6)
insert into XAPHUONG values(81,N'Đức Thuận',6)
insert into XAPHUONG values(82,N'Suối Khiết',6)

insert into XAPHUONG values(83,N'Võ Xu',7)
insert into XAPHUONG values(84,N'Đức Tài',7)
insert into XAPHUONG values(85,N'Đa Kai',7)
insert into XAPHUONG values(86,N'Sùng Nhơn',7)
insert into XAPHUONG values(87,N'Mê Pu',7)
insert into XAPHUONG values(88,N'Nam Chính',7)
insert into XAPHUONG values(89,N'Đức Hạnh',7)
insert into XAPHUONG values(90,N'Đức Tín',7)
insert into XAPHUONG values(91,N'Vũ Hòa',7)
insert into XAPHUONG values(92,N'Tân Hà',7)
insert into XAPHUONG values(93,N'Đông Hà',7)
insert into XAPHUONG values(94,N'Trà Tân',7)

insert into XAPHUONG values(95,N'Chợ Lầu',8)
insert into XAPHUONG values(96,N'Phan Sớn',8)
insert into XAPHUONG values(97,N'Phan Lâm',8)
insert into XAPHUONG values(98,N'Bình An',8)
insert into XAPHUONG values(99,N'Phan Điền',8)
insert into XAPHUONG values(100,N'Hải Ninh',8)
insert into XAPHUONG values(101,N'Sông Lũy',8)
insert into XAPHUONG values(102,N'Phan Tiến',8)
insert into XAPHUONG values(103,N'Sông Bình',8)
insert into XAPHUONG values(104,N'Lương Sơn',8)
insert into XAPHUONG values(105,N'Phan Hòa',8)
insert into XAPHUONG values(106,N'Phan Thanh',8)
insert into XAPHUONG values(107,N'Hồng Thái',8)
insert into XAPHUONG values(108,N'Phan Hiệp',8)
insert into XAPHUONG values(109,N'Bình Tân',8)
insert into XAPHUONG values(110,N'Phan Rí Thành',8)
insert into XAPHUONG values(111,N'Hòa Thắng',8)
insert into XAPHUONG values(112,N'Hồng Phong',8)

insert into XAPHUONG values(113,N'Phước Hội',9)
insert into XAPHUONG values(114,N'Phước Lộc',9)
insert into XAPHUONG values(115,N'Tân Thiện',9)
insert into XAPHUONG values(116,N'Tân An',9)
insert into XAPHUONG values(117,N'Bình Tân',9)
insert into XAPHUONG values(118,N'Tân Hải',9)
insert into XAPHUONG values(119,N'Tân Tiến',9)
insert into XAPHUONG values(120,N'Tân Bình',9)
insert into XAPHUONG values(121,N'Tân Phước',9)

insert into XAPHUONG values(122,N'Ngũ Phụng',10)
insert into XAPHUONG values(123,N'Long Hải',10)
insert into XAPHUONG values(124,N'Tam Thanh',10)

drop procedure report_by_huyen
create procedure report_by_huyen @ngay varchar(10)
as 
select ct.SO_CT,ct.NGAY_NHAN, ct.NGAY_HEN, ct.TEN_NGUOI_NHAN, ct.DIEN_THOAI,ct.DIA_CHI_THUONG_TRU ,ct.DIA_CHI , ct.HUYEN_NHAN , h.TEN_HUYENTP,ct.NGAY_DUYET,CT.CUOC_PHI from CT_HCC as ct, HUYEN_TP as h
where ct.HUYEN_NHAN=h.MA_HUYENTP and ct.Approved=1 and ct.NGAY_NHAN=@ngay
exec report_by_huyen_duyet @ngay='27/08/2022'

create procedure report_by_huyen_duyet @ngay varchar(10)
as 
select ct.SO_CT,ct.NGAY_NHAN, ct.NGAY_HEN, ct.TEN_NGUOI_NHAN, ct.DIEN_THOAI,ct.DIA_CHI_THUONG_TRU ,ct.DIA_CHI , ct.HUYEN_NHAN , h.TEN_HUYENTP,ct.NGAY_DUYET,CT.CUOC_PHI from CT_HCC as ct, HUYEN_TP as h
where ct.HUYEN_NHAN=h.MA_HUYENTP and ct.Approved=1 and ct.NGAY_DUYET=@ngay

create procedure report_by_huyen_hen @ngay varchar(10)
as 
select ct.SO_CT,ct.NGAY_NHAN, ct.NGAY_HEN, ct.TEN_NGUOI_NHAN, ct.DIEN_THOAI,ct.DIA_CHI_THUONG_TRU ,ct.DIA_CHI , ct.HUYEN_NHAN , h.TEN_HUYENTP,ct.NGAY_DUYET,CT.CUOC_PHI from CT_HCC as ct, HUYEN_TP as h
where ct.HUYEN_NHAN=h.MA_HUYENTP and ct.Approved=1 and ct.NGAY_HEN=@ngay

create procedure report_by_huyen_all
as 
select ct.SO_CT,ct.NGAY_NHAN, ct.NGAY_HEN, ct.TEN_NGUOI_NHAN, ct.DIEN_THOAI,ct.DIA_CHI_THUONG_TRU ,ct.DIA_CHI , ct.HUYEN_NHAN , h.TEN_HUYENTP,ct.NGAY_DUYET,CT.CUOC_PHI from CT_HCC as ct, HUYEN_TP as h
where ct.HUYEN_NHAN=h.MA_HUYENTP and ct.Approved=1

drop proc search_soct @ma=1
create proc search_soct @ma int
as
select * from CT_HCC
where SO_CT=@ma

create proc search_sdt @sdt varchar(50)
as
select * from CT_HCC
where DIEN_THOAI=@sdt

search_sdt '12312'

drop proc search_nguoinhan @ten='T'
create proc search_nguoinhan @ten nvarchar(50)
as
select * from CT_HCC
where TEN_NGUOI_NHAN=@ten


 
drop proc load_huyen_report
create proc load_huyen_report
as 
select MA_HUYENTP as mah, TEN_HUYENTP as tenhuyen
from HUYEN_TP

drop proc load_ct_huyen
create proc load_ct_huyen 
as
with mahuyen as
(
select MA_HUYENTP,TEN_HUYENTP from HUYEN_TP
)
--select MA_HUYENTP,TEN_HUYENTP,'','','','','','','' from mahuyen
--UNION ALL
select ct.SO_CT as "Số Seri",ct.NGAY_NHAN as "Ngày Nhận", ct.NGAY_HEN as "Ngày Hẹn", ct.TEN_NGUOI_NHAN as "Họ và tên", ct.DIEN_THOAI as "Số điện thoại",ct.DIA_CHI_THUONG_TRU as "Địa chỉ phát" , ct.DIA_CHI as "Địa chỉ trên hồ sơ" , ct.HUYEN_NHAN as Huyện,h.TEN_HUYENTP as "Tên Quận Huyện Nhận CT"
from CT_HCC as ct, HUYEN_TP as H, mahuyen as ma
where ct.HUYEN_NHAN=ma.MA_HUYENTP and ct.HUYEN_NHAN=h.MA_HUYENTP


exec load_ct_huyen

create proc load_ma_huyen
as
select MA_HUYENTP from HUYEN_TP

drop proc load_ct_huyen_theo_ma
create proc load_ct_huyen_theo_ma @huyen int
as
--select MA_HUYENTP,TEN_HUYENTP,'','','','','','','' from mahuyen
--UNION ALL
select ct.SO_CT as "Số Seri",ct.NGAY_NHAN as "Ngày Nhận", ct.NGAY_HEN as "Ngày Hẹn", ct.TEN_NGUOI_NHAN as "Họ và tên", ct.DIEN_THOAI as "Số điện thoại",ct.DIA_CHI_THUONG_TRU as "Địa chỉ phát" , ct.DIA_CHI as "Địa chỉ trên hồ sơ" , ct.HUYEN_NHAN as Huyện,h.TEN_HUYENTP as "Tên Quận Huyện Nhận CT"
from CT_HCC as ct, HUYEN_TP as H
where ct.HUYEN_NHAN=@huyen and ct.HUYEN_NHAN=h.MA_HUYENTP

exec load_ct_huyen_theo_ma @huyen=2
exec load_tinh_huyen_xa @tinh=86,@huyen=1,@xa=1


CREATE TYPE CT_HCC_Table AS TABLE
(
SO_CT	int PRIMARY KEY not null,
	Approved bit not null,
TEN_NGUOI_NHAN	nvarchar(50) not null ,	
DIA_CHI	nvarchar(50) not null,	
SO_HS_HCC	nvarchar(50) not null,	
NGAY_HEN	date	not null,
NGAY_NHAN	date	not null,
DIA_CHI_THUONG_TRU nvarchar(50) not null,
TINH_NHAN	int	not null,
HUYEN_NHAN	int	not null,
XA_NHAN	int	not null,
SO_HS_KEM	varchar(10)	not null,
DIEN_THOAI	varchar(50)	not null,
MA_LOAI_HS	int	not null,
TRONG_LUONG	varchar(10)not null,
MA_BUUGUI	varchar(10)not null	,
GHI_CHU	nvarchar(50)not null,
NGAY_DUYET DATE NULL,
CUOC_PHI decimal(20) NOT NULL
)


SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<m.nabeel siddiqui="">
-- Create date: <19/4/2016>
-- Description:	<adding records="" in="" moodleusers="">
-- =============================================
create PROCEDURE ImportTableExcel 
@table_excel CT_HCC_Table readonly
AS
BEGIN
insert into CT_HCC select * from @table_excel
END

drop proc search_tinh
create proc search_tinh @tinh nvarchar(50)
as
select T.TEN_TINH as "Tên Tỉnh" from TINH as T, CT_HCC as ct
where T.TEN_TINH=@tinh and ct.TINH_NHAN=T.MA_TINH

exec search_tinh @tinh= N'Bình Thuận'


create proc search_box_tinh @box nvarchar(50)
as
select TEN_TINH from TINH
where @box= TEN_TINH

search_box_tinh @box="Bình Thuận"

CREATE TABLE [dbo].[药品分类](
	[编号] [int] IDENTITY(1,1) NOT NULL,
	[名称] [varchar](200) NULL,
	[备注] [varchar](500) NULL)

CREATE TABLE [dbo].[药品出口单](
	[出口单号] [int] NULL,
	[药品编号] [int] NULL,
	[批号] [varchar](50) NULL,
	[数量] [varchar](50) NULL,
	[出口日期] [date] NULL,
	[验收人] [varchar](50) NULL

CREATE TABLE [dbo].[药典](
	[药品类别] [varchar](20) NULL,
	[流水号] [int] NULL,
	[通用名] [varchar](100) NULL,
	[速记码] [varchar](50) NULL,
	[生成企业(总代理商)] [varchar](200) NULL,
	[药库规格] [varchar](100) NULL,
	[剂型] [varchar](50) NULL,
	[采购价] [money] NULL,
	[ID] [int] IDENTITY(1,1) NOT NULL,

CREATE TABLE [dbo].[产地](
	[编号] [int] IDENTITY(1,1) NOT NULL,
	[产地名称] [varchar](100) NOT NULL,
	[备注] [varchar](200) NULL,

CREATE TABLE [dbo].[药品资料](
	[编号] [varchar](20) NOT NULL,
	[名称] [varchar](150) NOT NULL,
	[规格] [varchar](100) NULL,
	[整量单位] [varchar](50) NULL,
	[散量单位] [varchar](50) NULL,
	[入库单价] [decimal](12, 2) NULL,
	[出库单价] [decimal](12, 2) NOT NULL,
	[批发价] [decimal](12, 2) NULL,
	[整散比] [decimal](12, 2) NULL,
	[分类] [varchar](100) NULL,
	[费用归类] [varchar](100) NULL,
	[效期] [int] NOT NULL,
	[生产日期] [datetime] NULL,
	[上限] [decimal](12, 2) NULL,
	[下限] [decimal](12, 2) NULL,

CREATE TABLE [dbo].[药品信息单](
	[药品ID] [tinyint] NOT NULL,
	[药品名称] [varchar](50) NULL,
	[药品简码] [varchar](50) NULL,
	[俗名] [varchar](50) NULL,
	[俗名简码] [varchar](50) NULL,
	[药品类型] [varchar](50) NULL,
	[剂型] [varchar](50) NULL,
	[规格] [varchar](50) NULL,
	[批号] [varchar](50) NULL,
	[生产商] [varchar](50) NULL,
	[地址] [varchar](100) NULL,
	[是否报销品] [varchar](4) NULL,

CREATE TABLE [dbo].[药品入口单](
	[入库单号] [varchar](50) NULL,
	[药品编号] [varchar](50) NULL,
	[药品名称] [varchar](50) NULL,
	[批号] [varchar](50) NULL,
	[药品类型] [varchar](50) NULL,
	[规格] [varchar](50) NULL,
	[剂型] [varchar](50) NULL,
	[生产日期] [date] NULL,
	[有效日期] [date] NULL,
	[进价] [real] NULL,
	[售价] [real] NULL,
	[疾病功效] [varchar](50) NULL,
	[备注] [varchar](50) NULL,
	[数量] [varchar](50) NULL,
	[验收人] [varchar](50) NULL
)

CREATE TABLE [dbo].[药品库存信息表](
	[药品编号] [int] NULL,
	[药品名称] [varchar](50) NULL,
	[简码] [varchar](50) NULL,
	[俗名] [varchar](50) NULL,
	[规格] [varchar](50) NULL,
	[药品类型] [varchar](50) NULL,
	[剂型] [varchar](50) NULL,
	[批号] [varchar](50) NULL,
	[库存数量] [varchar](50) NULL,
	[库存日期] [date] NULL,
	[到期日期] [date] NULL,
	[库存位置] [varchar](50) NULL
) 
CREATE TABLE [dbo].[药品库存](
	[编号] [int] IDENTITY(1,1) NOT NULL,
	[库房] [varchar](20) NULL,
	[药品类型] [varchar](20) NULL,
	[药品编号] [varchar](20) NULL,
	[助记码] [varchar](50) NULL,
	[药品名] [varchar](100) NULL,
	[规格] [varchar](50) NULL,
	[单位] [varchar](50) NULL,
	[库存] [real] NULL,
	[批号] [varchar](20) NULL,
	[单价] [real] NULL,
	[用法] [varchar](50) NULL,
	[备注] [varchar](100) NULL,
	[状态] [varchar](10) NULL)	
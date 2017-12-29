create table table_name
(
  xls_id       VARCHAR2(10) not null, --主键需要发送的报表ID
  status       VARCHAR2(1),           --需要发送的报表状态 T表示开启
  scripts      VARCHAR2(4000),        --导出数据的SQL语句
  send_time    VARCHAR2(200),         --需要发送的报表时间 某一天的某个小时
  mail_subject VARCHAR2(500),         --需要发送的报表名称
  remark       VARCHAR2(500),         --需要发送的报表备注
  send_user    VARCHAR2(200),         --需要发送的报表邮箱
  save_add     VARCHAR2(100)          --导出数据的保存地址
)

alter table table_name
  add constraint table_name_pk primary key (XLS_ID)
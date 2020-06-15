create table cat (
  id                int not null auto_increment primary key,
  cat_name          varchar(100) not null,
  birthday_date     date
) ENGINE=InnoDB;
CREATE TABLE t_scan_dir(dirpath text, last_checked);

CREATE TABLE t_scan_file(fpath text, last_checked, wikiPageName text);

CREATE TABLE "t_scan_status" (base_dir, cur_dir, cur_file);

CREATE TABLE t_seq(wikiPageName int);

CREATE TABLE t_skip_file( 
  fpath text not null
  , constraint pk_skip_file primary key (fpath)
);

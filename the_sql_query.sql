/*
    Database    : Oracle 21c XE
    Author      : Shashi Kant Prasad
    Date        : 14 April 2022
    Description : SQL query that generates random sample data and imitates a table to some extent
    
    Dynamic Sample data generation

    Requirements:
        * SQL SELECT statement only.
        * Imitate a typical application table - wide ranging & adjustable change (row/record INSERT, UPDATE, DELETE) behaviour
        * Record structure:
            ** record_id: Primary Key
            ** group_fld: for "groub by" clause utility, cardinality [1, 4]
            ** string_fld: data field. Alpha-numberic, Length Range [4, 20]
            ** number_fld: data field. Range [-100, 100]
            ** date_fld: data field. Range [current_date - 7, current + 7]
        * Desired data characteristics:
            ** Arbitrary number of rows per run. Range [0, 50]
            ** Arbitrary number of rows differing every SQL run
            ** Unique ID field should be unique accross multiple runs of the SELECT statement

    Solution:
        * Recursive query.
        * Derive(deterministic function) dynamic fields from the unique id field.
*/
-- create or replace view c##xeread.table_sim_v as
with config(row_cnt_max, grp_cnt_max, var_percent, rr_size) as (
    select
        20, -- row_cnt_max  : Query execution time O(n(log(n))), n = row_cnt_max
        4,  -- grp_cnt_max  : as you wish
        40, -- var_percent  : Higher the number, higher the simulated DML(All) activity
        100 -- rr_size      : Random number range size. Higher the number, higher the simulated DML(Update only) activity
    from dual
),
str_lookup(key, val) as (
    -- Just some common English words
    select 0, 'not' from dual
    union select 1, 'his' from dual
    union select 2, 'from' from dual
    union select 3, 'they' from dual
    union select 4, 'we' from dual
    union select 5, 'her' from dual
    union select 6, 'an' from dual
    union select 7, 'my' from dual
    union select 8, 'their' from dual
    union select 9, 'who' from dual
    union select 10, 'him' from dual
    union select 11, 'your' from dual
    union select 12, 'could' from dual
    union select 13, 'them' from dual
    union select 14, 'than' from dual
    union select 15, 'its' from dual
    union select 16, 'also' from dual
    union select 17, 'two' from dual
    union select 18, 'our' from dual
    union select 19, 'any' from dual
    union select 20, 'these' from dual
    union select 21, 'us' from dual
),
run_consts(rr_dynamic, row_cnt_dynamic) as (
    -- derive dynamic quantites from input
    select
        rr_size * var_percent / 100, -- rr_dynamic
        trunc(DBMS_RANDOM.value(row_cnt_max * (100 - var_percent * 2 / 3) / 100, row_cnt_max)) -- row_cnt_dynamic
    from config
),
params(row_cnt_max, grp_cnt_max, var_percent, rr_size,
    rr_dynamic, row_cnt_dynamic, upd_fld_flag, str_lookup_size) as (
    -- form a record of all useful parameters
    select
        a.*,
        b.*,
        trunc(DBMS_RANDOM.value(0,  rr_dynamic)), -- upd_fld_flag
        (select count(*) from str_lookup) -- str_lookup_size
    from config a
        join run_consts b on (1 = 1)
),
key_set(id, sort_order) as (
    -- the recursive query to enumerate all primary keys
    select 1, DBMS_RANDOM.NORMAL() from dual
    union all
    select a.id + 1, DBMS_RANDOM.NORMAL()
    from key_set a
        join config b on (a.id < b.row_cnt_max)
),
key_set_shuffled(id) as (
    select id from key_set order by sort_order asc
),
key_subset(id, upd_row_select) as (
    select
        a.id,
        (case when b.var_percent > 0 and trunc(DBMS_RANDOM.value(0, b.rr_dynamic)) > trunc(DBMS_RANDOM.value(0, b.rr_size)) then 1 else 0 end)
    from key_set_shuffled a
        join params b on (1 = 1)
    where rownum <= b.row_cnt_dynamic
)
select -- finally! the output
    'R_' || lpad(a.id, 4, '0') as record_id,
    
    'G_' || lpad(1 + mod(a.id + a.upd_row_select * (case when b.upd_fld_flag > trunc(DBMS_RANDOM.value(0, b.rr_size)) then 1 else 0 end)  * b.grp_cnt_max / 2, b.grp_cnt_max), 2, '0')
        as group_fld,

    (select val from str_lookup where key = mod(a.id, b.str_lookup_size))
    || (select (case when mod(a.id, length(val)) = 0 then ' ' || val else '' end) from str_lookup where key = mod(b.row_cnt_max - a.id, b.str_lookup_size))
    || (case when a.upd_row_select = 1 and b.upd_fld_flag > trunc(DBMS_RANDOM.value(0, b.rr_size)) then ' CROSS' else '' end)
        as str_fld,

    mod(a.id - a.upd_row_select * (case when b.upd_fld_flag > trunc(DBMS_RANDOM.value(0, b.rr_size)) then 1 else 0 end) * b.row_cnt_max / 2, 100)
        as num_fld,
    
    mod(a.id - a.upd_row_select * (case when b.upd_fld_flag > trunc(DBMS_RANDOM.value(0, b.rr_size)) then 1 else 0 end) * 3, 7) + current_date
        as date_fld
from
    key_subset a
        join params b on (1 = 1);
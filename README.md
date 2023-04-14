# sql_dml_simulation
This is just a fun exercise.

<h2>About the SQL Query:</h2>

SQL "SELECT" simulates "INSERT", "UPDATE" &amp; "DELETE" statements, i.e., not using stored records in a database table but generating them in real-time.
Purpose: Dynamic Sample data generation. Different output on every run but not entirely random. Imitate a typical application table.

<pre>
Constraints:
    * SQL SELECT statement only. No DDL allowed.
    * Adjustable change (row/record INSERT, UPDATE, DELETE) behaviour.
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

Problem Solving Insights:
    * Recursive Query: To generate series of record IDs.
    * Hash Decoding: To derive data fields from record IDs. This ties data fields to record IDs.
    * Subset Selection: For arbitrary combination of records every run. Simulates INSERT/DELETE.
    * Dynamic Decision Parameter: To simulate UPDATE to random fields.
    * How to pick an arbitrary number of records from a set?
    * How to make the “variability” of data changes adjustable?

Note:
      * When the "var_percent" value is set 0, the query outputs the same records every run. The simulated DML activity increases as this value increases.
      * The query output may get corrupted for row count limits above 10000. It can be altered to accomodate large amounts of output.
</pre>

<h2>About Data Visualization:</h2>

The MS Excel file "Data Visualization.xlsm" (Macro-enabled Workbook) has several VBA macro procedures that connect to the database and fetch output of the aforementioned SQL query.

<pre>
The visualization process:
    * The Workbook maintains copies of the current and last runs of the SQL query.
    * Copy of current output of the query is compared with that old copy and a list of differences is computed.
    * The visualization is just these differences being applied to the old copy.
    * The old copy then becomes identical to the current copy.
    * This process repeats and we get a series of visible changes to data.
</pre>

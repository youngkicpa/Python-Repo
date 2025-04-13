import psycopg

with psycopg.connect("dbname=postgres user=postgres password=Ysk960826! host=localhost port=5432") as conn:
    with conn.cursor() as cur:
        # Execute multiple SQL commands separately
        cur.execute("""DROP TABLE IF EXISTS "tTable";""")
        cur.execute("""
            CREATE TEMP TABLE "tTable" (
                code INTEGER UNIQUE,
                name TEXT,
                freq INTEGER,
                pbalance Numeric DEFAULT 0,
                debit Numeric DEFAULT 0,
                credit Numeric DEFAULT 0
            );
        """)

        cur.execute("""
            INSERT INTO "tTable"(code, name)
            SELECT DISTINCT "actcode","actname"
            FROM "PriorFY24"
            ON CONFLICT (code) DO NOTHING;
        """)

        cur.execute("""
            INSERT INTO "tTable"(code, name)
            SELECT DISTINCT Segment4, ACCOUNT_NAME
            FROM "HanWhaFY24"
            ON CONFLICT (code) DO NOTHING;
        """)

        cur.execute("""
            CREATE TEMP VIEW "tView" AS 
            SELECT Segment4 AS 계정코드, count(Segment4) as 횟수, 
                   SUM(LINE_ACCOUNTED_DR) AS 차변, 
                   SUM(LINE_ACCOUNTED_CR) AS 대변 
            FROM "HanWhaFY24"
            GROUP BY 계정코드;
        """)

        cur.execute("""
            UPDATE "tTable"
            SET debit = t.차변, credit = t.대변, freq = t.횟수
            FROM "tView" t
            WHERE code = t.계정코드;
        """)

        cur.execute("""DROP VIEW IF EXISTS "tView";""")

        cur.execute("""
            UPDATE "tTable" 
            SET pbalance = amounts  
            FROM "PriorFY24"  
            WHERE code = actcode;
        """)

        # **Execute the final SELECT statement separately**
        cur.execute("""
            SELECT *, (pbalance + debit - credit) AS BALANCE 
            FROM "tTable"
            ORDER BY code;
        """)

        result = cur.fetchall()  # Fetch the actual results
        for row in result:
            print(row)

        conn.commit()  # Ensure changes are saved

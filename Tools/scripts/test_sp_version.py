"""
test_sp_version.py — test gọi usp_B20BOM_AutoVersion trực tiếp
"""
import pyodbc, json, os

cfg_path = os.path.join(os.path.dirname(__file__), 'config', 'db_config.json')
with open(cfg_path) as f:
    cfg = json.load(f)

conn_str = (
    f"DRIVER={{{cfg['driver']}}};"
    f"SERVER={cfg['server']};"
    f"DATABASE={cfg['database']};"
    f"UID={cfg['username']};"
    f"PWD={cfg['password']};"
    "TrustServerCertificate=yes;"
)
conn = pyodbc.connect(conn_str, autocommit=False)

SP = 'B10_Boho.dbo.usp_B20BOM_AutoVersion'
item_id  = 21407   # dùng giá trị thật từ SSMS
par_biz  = None
type_b   = None

print("=== Test 1: named params, None qua literal NULL ===")
try:
    sql = f"EXEC {SP} @_ItemId0=?, @_ParentBizDocId=NULL, @_TypeB=NULL"
    cur = conn.cursor()
    cur.execute(sql, [item_id])
    r = cur.fetchone()
    print("OK →", r[0] if r else None)
except Exception as e:
    print("FAIL:", e)

print("\n=== Test 2: tất cả literal NULL ===")
try:
    sql = f"EXEC {SP} @_ItemId0=NULL, @_ParentBizDocId=NULL, @_TypeB=NULL"
    cur = conn.cursor()
    cur.execute(sql)
    r = cur.fetchone()
    print("OK →", r[0] if r else None)
except Exception as e:
    print("FAIL:", e)

print("\n=== Test 3: dùng sp_executesql ===")
try:
    sql = (
        "EXEC sp_executesql "
        "N'EXEC B10_Boho.dbo.usp_B20BOM_AutoVersion @_ItemId0, @_ParentBizDocId, @_TypeB', "
        "N'@_ItemId0 VARCHAR(50), @_ParentBizDocId VARCHAR(50), @_TypeB VARCHAR(50)', "
        "@_ItemId0=?, @_ParentBizDocId=NULL, @_TypeB=NULL"
    )
    cur = conn.cursor()
    cur.execute(sql, [item_id])
    r = cur.fetchone()
    print("OK →", r[0] if r else None)
except Exception as e:
    print("FAIL:", e)

print("\n=== Test 4: positional params ===")
try:
    sql = f"EXEC {SP} ?, NULL, NULL"
    cur = conn.cursor()
    cur.execute(sql, [item_id])
    r = cur.fetchone()
    print("OK →", r[0] if r else None)
except Exception as e:
    print("FAIL:", e)

conn.close()

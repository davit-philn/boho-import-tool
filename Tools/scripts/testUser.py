import hashlib
import pyodbc

USERNAME = "USER001"        # user test
PASSWORD = "Boho@2026"  # điền vào đây

def try_hashes(pw):
    results = {}
    # Các cách encode phổ biến
    for enc in ["utf-8", "utf-16-le", "utf-16"]:
        b = pw.encode(enc)
        results[f"MD5({enc})"] = hashlib.md5(b).hexdigest().upper()
        results[f"MD5({enc}) x2"] = hashlib.md5(hashlib.md5(b).digest()).hexdigest().upper()
    return results

# Lấy hash từ DB
conn = pyodbc.connect("...")  # connection string của bạn
row = conn.execute(
    "SELECT Checksum, MD5Hash FROM SYS_User WHERE UserName=?", USERNAME
).fetchone()

checksum_hex = row[0].hex().upper() if row[0] else ""
md5hash_hex  = row[1].hex().upper() if row[1] else ""

print(f"Checksum  trong DB: {checksum_hex}")
print(f"MD5Hash   trong DB: {md5hash_hex}")
print()

for name, val in try_hashes(PASSWORD).items():
    match_cs = "✅ MATCH Checksum" if val == checksum_hex else ""
    match_md = "✅ MATCH MD5Hash"  if val == md5hash_hex  else ""
    print(f"{name:25s}: {val}  {match_cs}{match_md}")

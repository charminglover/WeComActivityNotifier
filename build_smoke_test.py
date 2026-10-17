"""Exercise Excel import, filtering, job generation, and image-table rendering."""
from pathlib import Path

from engine import DEFAULT_FIELDS, filter_records, make_jobs, read_excel, render_table


sample = Path(__file__).with_name("演示活动数据.xlsx")
fields, records = read_excel(str(sample))
assert records, "示例工作簿没有解析出活动记录"

first = records[0]
activity = str(first.get("活动名称") or "").strip()
filtered = filter_records(records, activity=activity) if activity else filter_records(records)
assert first in filtered, "活动条件筛选未保留匹配记录"

chosen_fields = [field for field in DEFAULT_FIELDS if field in fields][:4]
assert chosen_fields, "示例工作簿缺少可用于生成表格的字段"
jobs = make_jobs(filtered[:3], chosen_fields, page_size=8, columns_per_page=4)
image = render_table(jobs[0])
assert image.width > 0 and image.height > 0, "图片表格渲染失败"

print(f"Smoke test passed: {len(fields)} fields, {len(records)} records, image {image.width}x{image.height}.")

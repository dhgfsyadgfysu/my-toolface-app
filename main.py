import flet as ft
from io import BytesIO
import csv
from decimal import Decimal, InvalidOperation
import io
import re

# ==================== 核心解码函数 ====================
DEFAULT_EXCEL_COLUMNS = (
    'TOOLFACE_STEERING_VECTOR', 'TOOLFACE_ROTATING_VECTOR', 'TOOLFACE_TOTAL_VECTOR',
)
MISSING_TOKENS = {'', 'null', 'none', 'nan', 'na', 'n/a'}


def parse_packed(value):
    if value is None:
        return None
    if isinstance(value, bool):
        raise ValueError('布尔值不是打包数据')
    token = str(value).strip()
    if token.lower() in MISSING_TOKENS:
        return None
    try:
        if token.lower().startswith('0x'):
            number = int(token, 16)
        else:
            d = Decimal(token)
            if not d.is_finite() or d != d.to_integral_value():
                raise ValueError('打包值必须为有限整数')
            if not 0 <= d <= 65535:
                raise ValueError('打包值超出 uint16 范围')
            number = int(d)
    except (InvalidOperation, ValueError) as e:
        raise ValueError(f'无法读取打包整数 {token!r}: {e}') from e
    if not 0 <= number <= 65535:
        raise ValueError('打包值超出 uint16 范围')
    return number


def decode_toolface(value):
    if isinstance(value, bool) or not isinstance(value, int):
        raise TypeError('packed value must be an integer')
    if not 0 <= value <= 0x3FF:
        raise ValueError('expected a 10-bit packed value (0..1023)')
    a, s = (value >> 4) & 0x3F, value & 0x0F
    if a > 60 or s > 10:
        raise ValueError('code outside CCS angle 0..60 / strength 0..10')
    angle = a * 6
    return {'packed': value, 'angle_code': a, 'strength_code': s,
            'angle_deg_encoded': angle, 'angle_deg': angle % 360,
            'strength_percent': s * 10, 'direction_resolved': s != 0}


def decode_channel_record(record):
    if not isinstance(record, (bytes, bytearray)) or len(record) != 5:
        raise ValueError('expected exactly five bytes')
    channel = int.from_bytes(record[:2], 'little')
    if channel not in (11, 12, 13) or record[2] != 2:
        raise ValueError('expected channel 11/12/13 with size=2')
    r = decode_toolface(int.from_bytes(record[3:], 'little'))
    r['channel_id'] = channel
    r['phase'] = {11: 'rotating', 12: 'steering', 13: 'total'}[channel]
    return r


RESULT_FIELDS = ('packed', 'angle_code', 'strength_code', 'angle_deg_encoded',
                 'angle_deg', 'strength_percent', 'direction_resolved', 'status', 'error')
LABELS = {'ok': '编码合法', 'zero_strength': '零强度，方向不可靠',
          'missing': '缺失', 'invalid': '异常', 'repeated_header': '重复表头，未解码'}


def decode_value(value):
    r = dict.fromkeys(RESULT_FIELDS, '')
    try:
        p = parse_packed(value)
        if p is None:
            r['status'] = 'missing'
            return r
        r['packed'] = p
        r.update(decode_toolface(p))
        r['status'] = 'ok' if r['direction_resolved'] else 'zero_strength'
    except (TypeError, ValueError) as e:
        r['status'] = 'invalid'
        r['error'] = str(e)
    return r


# ==================== CSV 内存 I/O ====================
def decode_csv_bytes(file_bytes, columns, encoding='utf-8-sig', delimiter=','):
    if not columns or len(set(columns)) != len(columns):
        raise ValueError('必须指定互不重复的列名')
    src = csv.DictReader(BytesIO(file_bytes).read().decode(encoding).splitlines(), delimiter=delimiter)
    headers = src.fieldnames
    if not headers:
        raise ValueError('CSV 缺少表头')
    norm = [h.strip() for h in headers]
    lookup = dict(zip(norm, headers))
    miss = [c for c in columns if c not in lookup]
    if miss:
        raise ValueError(f'CSV 中没有这些列: {miss}')
    adds = [f'{c}__{f}' for c in columns for f in RESULT_FIELDS]
    buf = io.StringIO(newline='')
    w = csv.DictWriter(buf, fieldnames=headers + adds, delimiter=delimiter)
    w.writeheader()
    counts = {c: {'ok': 0, 'zero_strength': 0, 'missing': 0, 'invalid': 0} for c in columns}
    n = 0
    for row in src:
        for c in columns:
            r = decode_value(row[lookup[c]])
            counts[c][r['status']] += 1
            for f in RESULT_FIELDS:
                row[f'{c}__{f}'] = r[f]
        w.writerow(row)
        n += 1
    return buf.getvalue().encode('utf-8-sig'), n, counts


# ==================== Flet UI ====================
def main(page: ft.Page):
    page.title = 'CCS 等效工具面解码'
    page.theme_mode = ft.ThemeMode.LIGHT
    page.scroll = ft.ScrollMode.AUTO
    page.padding = 16

    picked_bytes = {'data': None, 'name': None}
    out_bytes = {'data': None, 'name': None, 'mime': None}

    status = ft.Text('')
    log = ft.Column(scroll=ft.ScrollMode.AUTO, height=240, spacing=4)

    save_btn = ft.Button('保存结果', icon=ft.Icons.SAVE, disabled=True)

    columns_tf = ft.TextField(
        label='解码列名（每行一个）',
        value='\n'.join(DEFAULT_EXCEL_COLUMNS),
        multiline=True, min_lines=3,
    )

    picker = ft.FilePicker()
    saver = ft.FilePicker()

    def get_columns():
        return [c.strip() for c in columns_tf.value.splitlines() if c.strip()]

    async def handle_pick_files(e):
        files = await picker.pick_files(
            allow_multiple=False,
            allowed_extensions=["csv"]  # 现在只支持 CSV
        )
        if files and len(files) > 0:
            f = files[0]
            try:
                with open(f.path, 'rb') as fh:
                    picked_bytes['data'] = fh.read()
                picked_bytes['name'] = f.name
                status.value = f'已选择：{f.name}（{len(picked_bytes["data"])} 字节）'
                save_btn.disabled = True
                out_bytes['data'] = None
            except Exception as ex:
                status.value = f'读取失败：{ex}'
            page.update()

    async def handle_save_file(e):
        if not out_bytes['data']:
            return
        path = await saver.save_file(
            file_name=out_bytes['name'],
            allowed_extensions=['csv'] # 只保存为 CSV
        )
        if path:
            try:
                with open(path, 'wb') as fh:
                    fh.write(out_bytes['data'])
                status.value = f'已保存：{path}'
            except Exception as ex:
                status.value = f'保存失败：{ex}'
            page.update()

    def run_decode(e):
        if not picked_bytes['data']:
            status.value = '请先选择文件'
            page.update()
            return
        cols = get_columns()
        if not cols:
            status.value = '请填写至少一个列名'
            page.update()
            return
        log.controls.clear()
        try:
            # 统一走 CSV 解码逻辑
            data, n, counts = decode_csv_bytes(picked_bytes['data'], cols)
            out_bytes['data'] = data
            out_bytes['name'] = picked_bytes['name'].rsplit('.', 1)[0] + '_decoded.csv'
            out_bytes['mime'] = 'csv'
            
            log.controls.append(ft.Text(f'✅ CSV 解码 {n} 行', weight=ft.FontWeight.BOLD))
            for c, cc in counts.items():
                log.controls.append(ft.Text(
                    f'   {c}: 正常{cc["ok"]} / 零强度{cc["zero_strength"]} / '
                    f'缺失{cc["missing"]} / 异常{cc["invalid"]}'))
            
            status.value = '解码完成，可点“保存结果”'
            save_btn.disabled = False
        except Exception as ex:
            status.value = f'解码失败：{ex}'
            save_btn.disabled = True
        page.update()

    save_btn.on_click = handle_save_file

    page.services.append(picker)
    page.services.append(saver)

    # 单值解码
    single_input = ft.TextField(label='打包值（空格/逗号/换行分隔，如 248 0xF8）',
                                multiline=True, min_lines=2)
    single_out = ft.Column(scroll=ft.ScrollMode.AUTO, height=200)

    def run_single(e):
        toks = [t for t in re.split(r'[\s,;，；]+', single_input.value.strip()) if t]
        single_out.controls.clear()
        if not toks:
            single_out.controls.append(ft.Text('没有输入'))
        for t in toks:
            r = decode_value(t)
            single_out.controls.append(ft.Text(
                f'{t} → 角度 {r["angle_deg"]}° / 强度 {r["strength_percent"]}% / '
                f'{LABELS.get(r["status"], r["status"])} {r.get("error", "")}'))
        page.update()

    # 通道记录
    rec_input = ft.TextField(label='5 字节通道记录，如 0D 00 02 F8 00')
    rec_out = ft.Text('')

    def run_record(e):
        try:
            r = decode_channel_record(bytes.fromhex(rec_input.value.replace(',', ' ')))
            rec_out.value = (f'通道 {r["channel_id"]}（{r["phase"]}）→ '
                             f'角度 {r["angle_deg"]}° / 强度 {r["strength_percent"]}%')
            rec_out.color = 'green'
        except Exception as ex:
            rec_out.value = f'失败：{ex}'
            rec_out.color = 'red'
        page.update()

    # 批量面板
    batch_panel = ft.Column([
        ft.Text('批量数据集解码', size=18, weight=ft.FontWeight.BOLD),
        ft.Text('⚠️ 注意：为了兼容手机端，请先在电脑上使用 Excel 将文件另存为 CSV 格式后再上传。', color='orange'),
        columns_tf,
        ft.Row([
            ft.Button('选择 CSV 文件', icon=ft.Icons.UPLOAD_FILE, on_click=handle_pick_files),
            ft.Button('开始解码', icon=ft.Icons.PLAY_ARROW, on_click=run_decode),
            save_btn,
        ], wrap=True),
        status,
        log,
    ], spacing=10)

    single_panel = ft.Column([
        ft.Text('单个/多个打包值', size=18, weight=ft.FontWeight.BOLD),
        single_input,
        ft.Button('解码', on_click=run_single),
        single_out,
    ], spacing=10)

    record_panel = ft.Column([
        ft.Text('通道记录解码', size=18, weight=ft.FontWeight.BOLD),
        rec_input,
        ft.Button('解码记录', on_click=run_record),
        rec_out,
    ], spacing=10)

    help_panel = ft.Column([
        ft.Text('使用说明', size=18, weight=ft.FontWeight.BOLD),
        ft.Text('• 编码：高6位方向码(0..60，×6°)，低4位强度码(0..10，×10%)。'),
        ft.Text('• 零强度时方向不可靠，状态为 zero_strength。'),
        ft.Text('• 编码合法 ≠ 来源正确 ≠ 测量新鲜。'),
        ft.Text('• 文件需为 CSV 格式，且列名须与上方列表一致。'),
        ft.Text('• 结果会在新文件中追加“列名__字段”，不会修改源文件。'),
    ], spacing=6)

    tab_bar = ft.TabBar(
        tabs=[
            ft.Tab(label='📂 批量'),
            ft.Tab(label='🔢 单值'),
            ft.Tab(label='🧩 通道'),
            ft.Tab(label='❓ 说明'),
        ]
    )

    tab_view = ft.TabBarView(
        expand=1,
        controls=[
            ft.Container(batch_panel, padding=10),
            ft.Container(single_panel, padding=10),
            ft.Container(record_panel, padding=10),
            ft.Container(help_panel, padding=10),
        ]
    )

    page.add(
        ft.Tabs(
            selected_index=0,
            length=4,
            expand=1,
            content=ft.Column(
                controls=[tab_bar, tab_view],
                expand=True,
            ),
        )
    )


if __name__ == '__main__':
    ft.run(main)
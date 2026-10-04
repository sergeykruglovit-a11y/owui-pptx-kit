---
name: pptx-restyle-universal
description: Универсальный перенос фирменного стиля между любыми PPTX без внешнего инструментария: агент сам анализирует XML и рендеры референса, пишет скрипты python-pptx под конкретную пару файлов, собирает результат клонированием элементов референса и проходит визуальный QA с независимым ревьюером-субагентом. Используй для «оформи X в стиле Y», «перенеси на шаблон», «сделай как в образце», когда скилл pptx-restyle недоступен или не справляется с нестандартным референсом.
---

# Универсальный restyle PPTX: методика качества

Каждый раз ты пишешь скрипты заново под конкретную пару файлов. Это нормально. Ненормально — пропускать этапы.
Качество обеспечивают не скрипты, а **порядок работы, ворота качества и визуальная проверка**.

Определение «хорошо»: человек, знающий референс, не отличит результат от слайдов, свёрстанных автором референса;
при этом содержание исходника перенесено полностью и без выдумок.

## Пять причин провала (и правило против каждой)

1. **Слепота.** Модель не смотрит на слайды → придумывает палитру и формат. Правило: смотри каждую картинку через
   `read_file(path)` — ТОЛЬКО параметр `path`, без `start_line/end_line` (пустые строки и 0 дают ошибку 422).
   `display_file` показывает файл пользователю, тебе он ничего не возвращает.
2. **Рисование с нуля.** Правило: каждый визуальный элемент результата — **копия** элемента референса (слайда,
   фигуры, макета). Свои цвета, шрифты, фигуры запрещены, кроме случая из §5.C.
3. **Скриншоты-подложки.** Рендер слайда референса фоном + плашки поверх — запрещено всегда.
4. **Утечки чужого контента.** Текст, цифры, фото, скриншоты интерфейсов из референса, оставшиеся на слайде.
   Правило: всё, что не заполнено новым содержанием и не является «хромом» (лого, колонтитул, номер), удаляется.
5. **Проверка «файл открывается».** Правило: «готово» — только после того, как каждая пара «шаблон | результат»
   просмотрена глазами и независимый ревьюер не нашёл дефектов уровня «критично».

## Роли (нативные субагенты OWUI, `delegate_task`)

- **Ты — лид.** Ведёшь задачи (`create_tasks`/`update_task`), принимаешь решения, пишешь план и сборку.
- **Аналитик референса** (субагент, по желанию, если в референсе > 12 слайдов или сложные макеты): строит паспорт
  стиля и каталог шаблонов (§3). Передай в `context` пути к файлам и требования к формату `work/ref/passport.md`.
- **Ревьюер** (субагент, ОБЯЗАТЕЛЬНО на этапе QA): свежий взгляд без твоих оправданий. Передай: путь к
  `work/qa/`, к `work/plan.json`, к `work/src/outline.md`, чек-лист §7 и требование вернуть список дефектов в формате
  `слайд N — [критично|важно|мелочь] — что не так — как исправить`. Ревьюер ничего не правит.
- Субагенты работают в том же терминале и видят те же файлы. Рекурсия запрещена; субагенты не делегируют дальше.
- Не дроби сборку между субагентами: один скрипт сборки — одно место правды.

## 0. Подготовка (ворота: окружение готово)

Рабочая папка `~/<название задачи>/`, всё промежуточное — в `work/`. Проверка:
```bash
python3 -c "import pptx, lxml, PIL; print('pptx', pptx.__version__)"; which soffice pdftoppm pdftotext fc-match
```
Неподходящий формат исходника (ppt/odp/key) → `soffice --headless --convert-to pptx`.
Скопируй проверенные рецепты (§9) в `work/lib.py` через `write_file` и импортируй их. Скрипты храни файлами
(`work/01_profile.py`, `work/02_outline.py`, `work/03_build.py`, `work/04_qa.py`), запускай через `run_command`
и выводи только итог (`| tail -40`), не огромные дампы.

## 1. Рендер обоих файлов

`render()` из §9 → `work/ref/render/ref-NN.png`, `work/src/render/src-NN.png`; собери контакт-листы и посмотри их.
Сразу запиши: формат (16:9 / 4:3 / иной — по `slide_width/slide_height`, никогда «на глаз»), сколько слайдов,
визуальные типы слайдов.

## 2. Паспорт стиля `work/ref/passport.md` (ворота: паспорт написан по данным, а не по впечатлению)

Собери скриптом, а не на глаз:
- **Палитра фактическая**: частоты `srgbClr` в `ppt/slides/*.xml` (вес 3) + layouts/masters (вес 1). Цвета темы
  (`theme1.xml`) часто дефолтные офисные и в дизайне не используются — не путай.
- **Шрифты фактические**: частоты `<a:latin typeface>`. Проверь наличие в рендерере: `fc-match "<Font>"`; если
  подмена — запиши: рендеры будут шире/уже оригинала.
- **Сетка**: поля, позиция лого, позиция и стиль заголовка, фон/панели (координаты в долях слайда).
- **Хром**: элементы, повторяющиеся на ≥40 % слайдов в той же позиции (лого, колонтитул, номер, плашки).

## 3. Каталог шаблонов `work/ref/catalog.md`

Для КАЖДОГО слайда референса и КАЖДОГО макета (layout):
- тип (обложка, тезис, N карточек, список, цифры/KPI, картинка+текст, таблица, финал);
- слоты: `id`, роль (заголовок/подзаголовок/текст карточки/подпись/цифра), bbox, размер шрифта, стили абзацев
  (у карточки «заголовок + описание» это два разных стиля в одной фигуре!), ёмкость ≈ знаков;
  ёмкость ≈ (ширина_pt / (0,55·кегль)) × (высота_pt / (1,2·кегль));
- карточки: какая декоративная фигура (fill) содержит какие текст/иконку (по центрам bbox);
- списки с маркерами-картинками (маленькие картинки слева от многоабзацного текста — по одной на строку);
- картинки: фон (>85 % площади), иконка (<2 %), контент (фото/скриншот — это контент референса!);
- **слайды-картинки** (одна картинка на весь слайд, текст «запечён») — помечай «не использовать как шаблон»;
- сложные группы (>12 фигур — макеты интерфейса, схемы) — как единое целое.
Нарисуй на рендерах прямоугольники с id слотов (`annot-NN.png`) и посмотри их — так ты не перепутаешь id.
Пустые плейсхолдеры — тоже слоты: позиция и кегль наследуются layout → master (`txStyles`).

## 4. Содержание исходника `work/src/outline.md`

Порядок, заголовок, абзацы с уровнями, таблицы, картинки (выгрузи в `work/src/img/`), заметки докладчика.
Красный/выделенный текст-пометку («дописать», «уточнить») не переноси — задай вопрос пользователю.

## 5. План `work/plan.json` (ворота: каждый слайд исходника сопоставлен, ёмкость проверена)

Для каждого слайда исходника выбери стратегию:
- **A. Клон слайда референса** (по умолчанию): тип и число элементов совпадают; лишние карточки/пункты удаляются
  вместе с их фигурами, оставшиеся карточки ряда равномерно растягиваются на ширину ряда (вместе с содержимым).
- **B. Слайд из макета** (`add_slide(layout)` + заполнение плейсхолдеров): когда подходящего готового слайда нет,
  а макет есть.
- **C. Композиция**: нужного типа нет вовсе (таблица, график, 5 пунктов при шаблонах на 3). Берёшь клон/макет с
  нужной сеткой и добавляешь элементы, построенные **только из примитивов референса**: копия существующей
  фигуры-карточки (`copy.deepcopy` элемента + смена `a:off`), цвета и шрифты — только из паспорта. Таблица:
  шапка — акцентный цвет палитры, текст — основной цвет текста, шрифт — основной шрифт.
Правила раскладки:
- Сохраняй порядок и смысл; сокращать можно формулировки, не сущности (названия, цифры, имена).
- Укладывайся в ёмкость слота; пункт с маркером-картинкой — строго одна строка.
- Чередуй шаблоны, когда подходит несколько; 10 одинаковых подряд — плохо.
- Заголовок результата — в слоте-заголовке шаблона, а не в любом верхнем текстовом поле.
Покажи план себе таблицей «слайд исходника → стратегия/шаблон → что куда» и проверь ёмкости до сборки.

## 6. Сборка `work/03_build.py`

Схема: открыть КОПИЮ референса → для каждого пункта плана клон/макет (§9) → заполнение (`set_paragraphs`) →
удаление незаполненных текстов, лишних маркеров и контентных картинок референса → композиция (стратегия C) →
удалить исходные слайды референса (`delete_slide`) → сохранить
`<исходник> — в стиле <референс>.pptx`. Мастер, макеты, тема и шрифты остаются от референса автоматически.
Подводные камни (все встречались на практике):
- **Не заменяй `cSld`/`spTree` целиком** при клонировании: `slide.shapes` закэширован на старый `spTree`, и всё,
  добавленное потом через `shapes.add_*`, пропадёт без ошибок. Переноси детей (см. `clone_slide`).
- Переназначай связи `r:embed`/`r:id`/`r:link` (картинки, графики, медиа); связи layout и notes не копируй.
- Заменяя текст, сохраняй `pPr`, `rPr` первого run и `endParaRPr` шаблонного абзаца; не трогай `bodyPr`
  (кроме якоря: многоабзацный слот → `anchor="t"`, иначе короткий текст съезжает на иконку).
- Отбивка между заголовком и текстом в карточке бывает пустым абзацем или завершающим `<a:br>` — сохрани её.
- Не задавай шрифты/цвета/кегли вручную в клонированных слотах — они уже правильные.
- Заметки докладчика исходника → `slide.notes_slide.notes_text_frame.text`.

## 7. QA (ворота: 0 критичных, ≤3 итераций)

1. Автоматика `work/04_qa.py`: рендер результата; `pair()` «шаблон | результат» для каждого слайда
   (для макета — образец макета с тестовым текстом); проверки: размер слайда = референсу; число слайдов = плану;
   абзацы текста шаблона, оставшиеся в результате и не заказанные планом (утечка); переполнение
   (`words_outside`: слова текста, отрисованные вне рамки своей фигуры, > 10 % → дефект).
2. Сам посмотри каждую `pair-NN.jpg` через `read_file`.
3. Делегируй ревью субагенту (см. «Роли») с чек-листом:
   - нет текста/фото/скриншотов референса, в том числе внутри картинок; нет пустых карточек и маркеров без текста;
   - текст не выходит за карточки, не налезает на соседей и декор;
   - сетка, позиция и стиль заголовка, лого — как в шаблоне; нет больших пустот там, где в шаблоне был контент;
   - палитра и шрифты — только из паспорта; формат как у референса;
   - содержание слайда соответствует исходнику (сверка с outline.md), ничего не потеряно и не выдумано.
4. Исправь план/сборку, пересобери, повтори QA. После 3 итераций — остановись и честно перечисли остаточные
   компромиссы.
Если шрифт референса подменён в рендерере, лёгкие отличия переносов допустимы; явное переполнение — нет.

## 8. Сдача

`display_file` итогового .pptx и контакт-листа результата. Отчёт: число слайдов; какие шаблоны/стратегии;
что сокращено/объединено; компромиссы; вопросы (пометки редактора, недостающие данные); что улучшит результат
(установить шрифт референса в `~/.fonts` + `fc-cache -f`). Слова «точное соответствие» — только если ревьюер
не нашёл ни одного дефекта.

## 9. Проверенные рецепты (скопируй в `work/lib.py`)

```python
# Проверенные рецепты для скилла pptx-restyle-universal. Скопируй в work/lib.py и импортируй из своих скриптов.
import copy, re, shutil, subprocess, tempfile
from pathlib import Path
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
SHAPES = {P + 'sp', P + 'pic', P + 'grpSp', P + 'graphicFrame', P + 'cxnSp'}


def ptext(p):  # текст абзаца, <a:br> -> \n
    return ''.join(e.text or '' if e.tag == A + 't' else '\n' if e.tag == A + 'br' else '' for e in p.iter())


def xfrm(el):
    """(x, y, cx, cy, chx, chy, chcx, chcy) из spPr/grpSpPr/graphicFrame; None если позиция наследуется."""
    pr = el.find(P + 'spPr')
    if pr is None:
        pr = el.find(P + 'grpSpPr')
    xf = pr.find(A + 'xfrm') if pr is not None else el.find(P + 'xfrm')
    if xf is None or xf.find(A + 'off') is None:
        return None
    o, e = xf.find(A + 'off'), xf.find(A + 'ext')
    v = [int(o.get('x')), int(o.get('y')), int(e.get('cx')), int(e.get('cy'))]
    co, ce = xf.find(A + 'chOff'), xf.find(A + 'chExt')
    return v + ([int(co.get('x')), int(co.get('y')), int(ce.get('cx')), int(ce.get('cy'))] if co is not None else v)


def walk(tree, tf=(0, 0, 1.0, 1.0), parent=None):
    """Все фигуры (включая вложенные в группы) с АБСОЛЮТНЫМ bbox в EMU — координаты детей группы пересчитаны."""
    for el in tree:
        if el.tag not in SHAPES:
            continue
        nv = el.find('.//' + P + 'cNvPr')
        x = xfrm(el)
        ox, oy, sx, sy = tf
        bbox = (ox + x[0] * sx, oy + x[1] * sy, x[2] * sx, x[3] * sy) if x else None
        yield {'el': el, 'id': nv.get('id'), 'name': nv.get('name'), 'tag': el.tag.split('}')[1], 'bbox': bbox, 'parent': parent}
        if el.tag == P + 'grpSp' and x:
            ksx, ksy = (x[2] / x[6] if x[6] else 1), (x[3] / x[7] if x[7] else 1)
            yield from walk(el, (ox + x[0] * sx - x[4] * ksx * sx, oy + x[1] * sy - x[5] * ksy * sy, ksx * sx, ksy * sy), nv.get('id'))


def by_id(slide, sid):
    for s in walk(slide._element.cSld.spTree):
        if s['id'] == str(sid):
            return s['el']


def clone_slide(prs, src):
    """Копия слайда src в конец prs. ВАЖНО: не заменять cSld/spTree целиком — python-pptx кэширует
    slide.shapes на старый spTree, и всё, что потом добавишь через shapes.add_*, уйдёт в никуда."""
    new = prs.slides.add_slide(src.slide_layout)
    sp_new = new._element.cSld.spTree
    for ch in list(sp_new):
        sp_new.remove(ch)
    src_cSld = copy.deepcopy(src._element.cSld)
    for ch in list(src_cSld.spTree):
        sp_new.append(ch)
    bg = src_cSld.find(P + 'bg')
    if bg is not None:
        new._element.cSld.insert(0, bg)
    rmap = {}
    for rel in list(src.part.rels.values()):
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue
        rmap[rel.rId] = (new.part.relate_to(rel.target_ref, rel.reltype, is_external=True) if rel.is_external
                         else new.part.relate_to(rel.target_part, rel.reltype))
    for el in new._element.iter():
        for k, v in list(el.attrib.items()):
            if k.startswith(R) and v in rmap:
                el.set(k, rmap[v])
    return new


def delete_slide(prs, slide):
    lst = prs.slides._sldIdLst
    for sid in list(lst):
        if prs.part.related_part(sid.rId) is slide.part:
            rid = sid.rId
            lst.remove(sid)
            prs.part.drop_rel(rid)
            return


def set_paragraphs(el, items):
    """Заменить текст фигуры, СОХРАНИВ форматирование шаблона. items: список строк (абзацев).
    Абзац i получает стиль i-го НЕПУСТОГО абзаца шаблона (последний — для остальных); пустые абзацы-отбивки
    между стилями шаблона сохраняются; завершающий <a:br> шаблонного абзаца сохраняется; '<br>' внутри строки — перенос."""
    tb = el.find(P + 'txBody')
    tps = tb.findall(A + 'p')
    full = [i for i, p in enumerate(tps) if ptext(p).strip()] or [0]
    out = []
    for j, text in enumerate(items):
        k = full[min(j, len(full) - 1)]
        if 0 < j < len(full):
            out += [copy.deepcopy(tps[i]) for i in range(full[j - 1] + 1, full[j])]
        tp = tps[k]
        rpr = next((r.find(A + 'rPr') for r in tp.iter(A + 'r') if r.find(A + 'rPr') is not None), None)
        if rpr is None:
            rpr = tp.find(A + 'endParaRPr')
        p = copy.deepcopy(tp)
        end = p.find(A + 'endParaRPr')
        for ch in list(p):
            if ch.tag != A + 'pPr':
                p.remove(ch)
        for li, line in enumerate(re.split(r'<br\s*/?>', text)):
            if li:
                p.append(etree.Element(A + 'br'))
            r = etree.SubElement(p, A + 'r')
            if rpr is not None:
                rp = copy.deepcopy(rpr)
                rp.tag = A + 'rPr'
                r.append(rp)
            etree.SubElement(r, A + 't').text = line
        if ptext(tp).endswith('\n'):
            p.append(etree.Element(A + 'br'))
        if end is not None:
            p.append(end)
        out.append(p)
    if len(full) > 1:  # многоабзацный слот: якорь вверх, чтобы короткий текст не «уезжал» на иконку
        bp = tb.find(A + 'bodyPr')
        if bp is not None and bp.get('anchor') in ('ctr', 'b'):
            bp.set('anchor', 't')
    for p in tps:
        tb.remove(p)
    for p in out:
        tb.append(p)


def render(pptx, outdir, prefix, dpi=80):
    """pptx -> outdir/prefix.pdf + prefix-01.png …  (отдельный профиль LibreOffice, чтобы не ловить lock)."""
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        shutil.copy(pptx, f'{td}/in.pptx')
        subprocess.run(['soffice', f'-env:UserInstallation=file://{td}/lo', '--headless', '--convert-to', 'pdf',
                        '--outdir', td, f'{td}/in.pptx'], capture_output=True, timeout=600)
        shutil.copy(f'{td}/in.pdf', outdir / f'{prefix}.pdf')
    for f in outdir.glob(f'{prefix}-*.png'):
        f.unlink()
    subprocess.run(['pdftoppm', '-r', str(dpi), '-png', str(outdir / f'{prefix}.pdf'), str(outdir / f'{prefix}-t')], timeout=600)
    pngs = sorted(outdir.glob(f'{prefix}-t-*.png'), key=lambda p: int(re.findall(r'(\d+)\.png$', p.name)[0]))
    res = []
    for i, p in enumerate(pngs, 1):
        dst = outdir / f'{prefix}-{i:02d}.png'
        p.rename(dst)
        res.append(dst)
    return res


def pair(left, right, out, lt='TEMPLATE', rt='RESULT', w=760):
    from PIL import Image, ImageDraw
    a, b = Image.open(left).convert('RGB'), Image.open(right).convert('RGB')
    a, b = a.resize((w, int(a.height * w / a.width))), b.resize((w, int(b.height * w / b.width)))
    im = Image.new('RGB', (2 * w + 30, max(a.height, b.height) + 28), (235, 235, 235))
    d = ImageDraw.Draw(im)
    d.text((10, 6), lt, fill='black'); d.text((w + 20, 6), rt, fill='black')
    im.paste(a, (10, 22)); im.paste(b, (w + 20, 22))
    im.save(out, quality=85)


def words_outside(pdf, page_no, slide_w, bbox, text, tol=0.015):
    """Сколько слов текста фигуры отрендерилось вне её bbox (переполнение). page_no с 1."""
    out = subprocess.run(['pdftotext', '-f', str(page_no), '-l', str(page_no), '-bbox', str(pdf), '-'],
                         capture_output=True, text=True).stdout
    pw = float(re.search(r'<page width="([\d.]+)" height="([\d.]+)"', out).group(1))
    ph = float(re.search(r'<page width="([\d.]+)" height="([\d.]+)"', out).group(2))
    k = pw / slide_w
    x0, y0, x1, y1 = bbox[0] * k, bbox[1] * k, (bbox[0] + bbox[2]) * k, (bbox[1] + bbox[3]) * k
    pos = {}
    for m in re.finditer(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</word>', out):
        pos.setdefault(m.group(5).lower().strip('.,:;!?«»"()'), []).append(
            ((float(m.group(1)) + float(m.group(3))) / 2, (float(m.group(2)) + float(m.group(4))) / 2))
    from collections import Counter
    need = Counter(w.lower().strip('.,:;!?«»"()') for w in text.split() if len(w) >= 4)
    t = tol * ph
    missing = total = 0
    for w, n in need.items():
        if w not in pos:
            continue
        total += n
        inside = sum(1 for cx, cy in pos[w] if x0 - t <= cx <= x1 + t and y0 - t <= cy <= y1 + t)
        missing += max(0, n - inside)
    return missing, total   # переполнение, если missing / total > 0.1
```

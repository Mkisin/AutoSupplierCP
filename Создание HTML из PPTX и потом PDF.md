Ниже — рабочий план создания **HTML-лонгрида, максимально точно повторяющего PPTX-презентацию**, но при этом остающегося редактируемым шаблоном с текстами, изображениями, таблицами и другими данными в виде переменных.

## 1. Сначала определить формат результата

Рекомендую строить страницу в двух режимах:

1. **Режим презентации** — каждый экран соответствует одному слайду, сохраняет исходное соотношение сторон и максимально точно повторяет PPTX.
2. **Режим лонгрида** — слайды идут вертикально один за другим и масштабируются под ширину экрана.

Каждый слайд представляется отдельной HTML-секцией:

```html
<section class="slide slide--cover">
  ...
</section>

<section class="slide slide--statistics">
  ...
</section>
```

Важно заранее принять решение: приоритетом будет точность или адаптивность.

Для максимально точной копии PPTX лучше использовать:

```css
.slide {
  width: 1920px;
  height: 1080px;
  position: relative;
}
```

Затем весь слайд пропорционально масштабировать под экран. Не нужно перестраивать его элементы на мобильных устройствах — иначе дизайн начнет отличаться от презентации.

---

# 2. Подготовить исходную презентацию

Для начала понадобится сама презентация `.pptx`.

Перед переносом нужно определить:

* размер слайда;
* соотношение сторон, обычно 16:9;
* используемые шрифты;
* основные цвета;
* фон каждого типа слайдов;
* повторяющиеся элементы;
* сетку и отступы;
* расположение текстов;
* размеры заголовков;
* формы, линии, иконки;
* изображения и способы их обрезки;
* тени, прозрачности и скругления.

Также желательно разделить слайды на типы, например:

```text
01. Обложка
02. Раздел
03. Текст + изображение
04. Две колонки
05. Карточки
06. Диаграмма
07. Цитата
08. Финальный слайд
```

Если в презентации 30 слайдов, это не обязательно означает 30 уникальных шаблонов. Обычно достаточно создать 5–10 типов слайдов и передавать в них разные данные.

---

# 3. Извлечь ресурсы из PPTX

Файл `.pptx` технически является ZIP-архивом.

Его можно распаковать и получить:

```text
ppt/
├── media/
├── slides/
├── slideLayouts/
├── slideMasters/
├── theme/
└── presentation.xml
```

Из папки `ppt/media` можно достать:

* фотографии;
* иллюстрации;
* SVG;
* PNG;
* JPEG;
* видео;
* другие встроенные файлы.

Можно переименовать:

```text
presentation.pptx
```

в:

```text
presentation.zip
```

и распаковать архив.

Однако XML внутри PPTX неудобно читать вручную. Для автоматического анализа лучше использовать Python:

```bash
pip install python-pptx
```

Пример получения размеров и координат объектов:

```python
from pptx import Presentation

prs = Presentation("presentation.pptx")

print("Ширина:", prs.slide_width)
print("Высота:", prs.slide_height)

for slide_index, slide in enumerate(prs.slides):
    print(f"Слайд {slide_index + 1}")

    for shape in slide.shapes:
        print({
            "name": shape.name,
            "left": shape.left,
            "top": shape.top,
            "width": shape.width,
            "height": shape.height,
            "type": shape.shape_type,
        })
```

Так можно получить базовую структуру слайда и использовать координаты объектов при создании CSS.

---

# 4. Сделать эталонные изображения слайдов

Каждый слайд нужно экспортировать в PNG или JPEG.

Эти изображения станут визуальными эталонами:

```text
references/
├── slide-01.png
├── slide-02.png
├── slide-03.png
└── ...
```

Экспортировать можно через PowerPoint:

```text
Файл → Экспорт → PNG
```

Желательно экспортировать слайды в высоком разрешении, например:

```text
1920 × 1080
```

или:

```text
2560 × 1440
```

Затем HTML-слайд можно сравнивать с эталонным изображением наложением или переключением прозрачности.

---

# 5. Выбрать технологический стек

Для такого проекта есть два разумных варианта.

## Вариант A. HTML + CSS + JavaScript

Подходит, если:

* презентация относительно небольшая;
* данные меняются редко;
* шаблонизация простая;
* страницу будет редактировать разработчик.

Структура:

```text
project/
├── index.html
├── css/
│   ├── reset.css
│   ├── tokens.css
│   ├── slides.css
│   └── responsive.css
├── js/
│   ├── data.js
│   ├── templates.js
│   └── app.js
├── assets/
│   ├── fonts/
│   ├── images/
│   └── icons/
└── references/
```

## Вариант B. React или Vue

Подходит, если:

* шаблон будет часто использоваться повторно;
* слайды нужно собирать из компонентов;
* данные будут поступать из JSON или API;
* понадобится несколько презентаций на одном шаблоне;
* нужны интерактивность и удобное масштабирование проекта.

Пример компонентов:

```text
src/
├── components/
│   ├── SlideFrame.jsx
│   ├── CoverSlide.jsx
│   ├── TextImageSlide.jsx
│   ├── CardsSlide.jsx
│   └── QuoteSlide.jsx
├── data/
│   └── presentation.json
├── styles/
│   ├── tokens.css
│   └── slides.css
└── App.jsx
```

Для первой версии я рекомендую **Vite + React + обычный CSS**. Это даст удобную шаблонизацию без лишней сложности.

---

# 6. Создать дизайн-токены

Все повторяющиеся параметры презентации нужно вынести в CSS-переменные.

```css
:root {
  --slide-width: 1920;
  --slide-height: 1080;

  --color-background: #f4f0e8;
  --color-surface: #ffffff;
  --color-text: #1d1d1b;
  --color-muted: #77746e;
  --color-accent: #e75b3b;
  --color-secondary: #243d4f;

  --font-heading: "Montserrat", Arial, sans-serif;
  --font-body: "Inter", Arial, sans-serif;

  --font-size-display: 92px;
  --font-size-h1: 72px;
  --font-size-h2: 48px;
  --font-size-body: 28px;
  --font-size-caption: 20px;

  --radius-card: 24px;
  --shadow-card: 0 20px 60px rgb(0 0 0 / 12%);
}
```

Это позволит изменить стиль всей презентации в одном месте.

Например, если нужно заменить акцентный цвет, достаточно изменить:

```css
--color-accent: #e75b3b;
```

---

# 7. Создать базовый компонент слайда

Каждый слайд должен иметь одинаковую внутреннюю систему координат.

Пример для React:

```jsx
export function SlideFrame({ className = "", children }) {
  return (
    <section className={`slide ${className}`}>
      <div className="slide__canvas">
        {children}
      </div>
    </section>
  );
}
```

CSS:

```css
.slide {
  display: flex;
  justify-content: center;
  width: 100%;
  overflow: hidden;
}

.slide__canvas {
  position: relative;
  flex: 0 0 auto;
  width: 1920px;
  height: 1080px;
  overflow: hidden;
  background: var(--color-background);
  transform-origin: top center;
}
```

Масштаб рассчитывается через JavaScript:

```js
function updateSlideScale() {
  const slideWidth = 1920;
  const viewportWidth = window.innerWidth;
  const scale = Math.min(viewportWidth / slideWidth, 1);

  document.documentElement.style.setProperty(
    "--slide-scale",
    scale
  );
}

window.addEventListener("resize", updateSlideScale);
updateSlideScale();
```

```css
.slide {
  height: calc(1080px * var(--slide-scale));
}

.slide__canvas {
  transform: scale(var(--slide-scale));
}
```

Таким образом, внутренняя верстка всегда остается в координатах `1920 × 1080`, но на экране масштабируется пропорционально.

---

# 8. Перенести координаты из PPTX в CSS

В PowerPoint элементы имеют координаты относительно слайда.

В HTML для точной копии можно использовать абсолютное позиционирование:

```css
.slide-cover__title {
  position: absolute;
  left: 128px;
  top: 194px;
  width: 1020px;
  height: 280px;
}
```

Компонент:

```jsx
function CoverSlide({ title, subtitle, image }) {
  return (
    <SlideFrame className="slide-cover">
      <h1 className="slide-cover__title">{title}</h1>

      <p className="slide-cover__subtitle">
        {subtitle}
      </p>

      <img
        className="slide-cover__image"
        src={image}
        alt=""
      />
    </SlideFrame>
  );
}
```

Для максимального совпадения не стоит пытаться сразу построить универсальную адаптивную сетку. Сначала нужно получить точную десктопную копию, а потом добавлять адаптацию.

---

# 9. Сделать данные отдельно от верстки

Основной принцип шаблона:

> Компоненты отвечают за дизайн, JSON — за содержание.

Пример файла `presentation.json`:

```json
{
  "meta": {
    "title": "Название презентации",
    "author": "Компания"
  },
  "slides": [
    {
      "id": "cover",
      "type": "cover",
      "title": "Главный заголовок презентации",
      "subtitle": "Краткое описание",
      "image": "/assets/images/cover.webp"
    },
    {
      "id": "about",
      "type": "text-image",
      "eyebrow": "О компании",
      "title": "Заголовок второго слайда",
      "text": "Основной текст второго слайда.",
      "image": "/assets/images/about.webp",
      "imagePosition": "right"
    },
    {
      "id": "advantages",
      "type": "cards",
      "title": "Наши преимущества",
      "items": [
        {
          "title": "Преимущество 1",
          "text": "Описание преимущества"
        },
        {
          "title": "Преимущество 2",
          "text": "Описание преимущества"
        }
      ]
    }
  ]
}
```

Рендеринг:

```jsx
function SlideRenderer({ slide }) {
  switch (slide.type) {
    case "cover":
      return <CoverSlide {...slide} />;

    case "text-image":
      return <TextImageSlide {...slide} />;

    case "cards":
      return <CardsSlide {...slide} />;

    default:
      return null;
  }
}
```

```jsx
function Presentation({ data }) {
  return (
    <main>
      {data.slides.map((slide) => (
        <SlideRenderer key={slide.id} slide={slide} />
      ))}
    </main>
  );
}
```

---

# 10. Использовать переменные для текста и изображений

Есть два уровня шаблонизации.

## Уровень 1. JSON

Самый удобный вариант для большинства проектов:

```json
{
  "title": "Новый заголовок",
  "description": "Новый текст",
  "image": "/uploads/new-image.jpg"
}
```

## Уровень 2. Плейсхолдеры

Можно использовать обозначения:

```text
{{company_name}}
{{slide_1_title}}
{{slide_1_image}}
{{report_date}}
```

Например:

```json
{
  "company_name": "ООО «Компания»",
  "report_date": "Июль 2026",
  "slide_1_title": "Годовой отчет",
  "slide_1_image": "/assets/images/report-cover.jpg"
}
```

Шаблон:

```html
<h1>{{slide_1_title}}</h1>
<p>{{company_name}}</p>
```

Для обработки таких переменных можно использовать:

* Handlebars;
* Mustache;
* Nunjucks;
* Liquid;
* собственную небольшую функцию;
* React props.

Пример с Handlebars:

```html
<script id="cover-template" type="text/x-handlebars-template">
  <section class="slide">
    <h1>{{title}}</h1>
    <p>{{subtitle}}</p>
    <img src="{{image}}" alt="">
  </section>
</script>
```

Для React отдельный шаблонизатор обычно не нужен: его роль выполняют props и JSX.

---

# 11. Настроить изображения как в PowerPoint

В презентациях изображения часто помещаются в рамки с обрезкой.

В HTML это воспроизводится через `object-fit`.

```css
.image-frame {
  overflow: hidden;
  border-radius: 24px;
}

.image-frame img {
  display: block;
  width: 100%;
  height: 100%;
  object-fit: cover;
  object-position: 50% 50%;
}
```

Положение изображения тоже можно передавать через данные:

```json
{
  "image": "/assets/images/person.jpg",
  "imagePosition": "40% 25%"
}
```

```jsx
<img
  src={image}
  alt=""
  style={{ objectPosition: imagePosition }}
/>
```

Для разных вариантов можно предусмотреть:

```json
{
  "imageFit": "cover",
  "imagePosition": "center top",
  "imageScale": 1.1
}
```

Это приблизит работу с изображением к механике PowerPoint.

---

# 12. Подключить исходные шрифты

Сначала нужно проверить лицензию шрифта. Если веб-использование разрешено, подключить локальные файлы через `@font-face`.

```css
@font-face {
  font-family: "Presentation Font";
  src:
    url("../assets/fonts/presentation-font.woff2") format("woff2");
  font-weight: 400;
  font-style: normal;
  font-display: swap;
}
```

Если исходный шрифт нельзя использовать на сайте, придется подобрать максимально близкую замену.

Необходимо учитывать, что даже визуально похожий шрифт может иметь другую ширину символов. Из-за этого:

* меняется перенос строк;
* заголовок занимает другую высоту;
* блоки перестают совпадать;
* элементы сдвигаются.

Поэтому шрифты — один из первых параметров, который нужно зафиксировать.

---

# 13. Воспроизвести основные элементы PowerPoint

Большинство элементов можно реализовать обычным CSS.

## Прямоугольники и карточки

```css
.card {
  background: #ffffff;
  border-radius: 28px;
  box-shadow: 0 24px 70px rgb(0 0 0 / 10%);
}
```

## Круги

```css
.circle {
  width: 240px;
  height: 240px;
  border-radius: 50%;
}
```

## Градиенты

```css
.gradient {
  background:
    linear-gradient(
      135deg,
      #243d4f 0%,
      #13222d 100%
    );
}
```

## Полупрозрачные фигуры

```css
.overlay {
  background: rgb(255 255 255 / 20%);
  backdrop-filter: blur(20px);
}
```

## Линии

```css
.line {
  height: 2px;
  background: currentColor;
}
```

## Произвольные формы

Использовать SVG:

```html
<svg viewBox="0 0 500 300">
  <path d="..." fill="currentColor" />
</svg>
```

Сложные декоративные фигуры из PowerPoint лучше экспортировать в SVG, а не пытаться повторять множеством HTML-элементов.

---

# 14. Создать ограничители для переменных текстов

Главная проблема шаблонизации: новый текст может оказаться длиннее исходного.

Для каждого текстового поля нужно определить правила:

```json
{
  "title": {
    "maxCharacters": 70,
    "maxLines": 3
  }
}
```

CSS:

```css
.title {
  display: -webkit-box;
  overflow: hidden;
  -webkit-box-orient: vertical;
  -webkit-line-clamp: 3;
}
```

Но простого обрезания часто недостаточно. Можно автоматически уменьшать шрифт до тех пор, пока текст не поместится.

Пример функции:

```js
function fitText(element, options = {}) {
  const {
    minFontSize = 24,
    step = 1
  } = options;

  let fontSize = parseFloat(
    getComputedStyle(element).fontSize
  );

  while (
    element.scrollHeight > element.clientHeight &&
    fontSize > minFontSize
  ) {
    fontSize -= step;
    element.style.fontSize = `${fontSize}px`;
  }
}
```

Использование:

```js
document
  .querySelectorAll("[data-fit-text]")
  .forEach((element) => {
    fitText(element, {
      minFontSize: Number(element.dataset.minFontSize || 24)
    });
  });
```

Однако лучше дополнительно задавать редакторские ограничения:

```text
Заголовок: до 70 знаков
Подзаголовок: до 160 знаков
Описание карточки: до 220 знаков
Количество карточек: от 3 до 5
```

---

# 15. Сделать компоненты с вариантами

Не стоит создавать отдельный компонент на каждую небольшую вариацию.

Например:

```jsx
<TextImageSlide
  imagePosition="right"
  theme="light"
/>
```

или:

```jsx
<TextImageSlide
  imagePosition="left"
  theme="dark"
/>
```

В JSON:

```json
{
  "type": "text-image",
  "variant": "image-left",
  "theme": "dark"
}
```

CSS:

```css
.slide-text-image--image-left .slide-text-image__content {
  left: 1020px;
}

.slide-text-image--image-left .slide-text-image__image {
  left: 0;
}

.slide--dark {
  color: #ffffff;
  background: #172630;
}
```

Полезно предусмотреть:

* `theme`;
* `variant`;
* `imagePosition`;
* `alignment`;
* `columns`;
* `accent`;
* `background`;
* `compact`;
* `showSlideNumber`.

---

# 16. Настроить точное визуальное сравнение

Чтобы проверить совпадение HTML с PPTX, удобно накладывать эталонный PNG поверх HTML.

```css
.slide-debug-reference {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  pointer-events: none;
  opacity: 0.5;
  z-index: 999;
}
```

```jsx
{debug && (
  <img
    className="slide-debug-reference"
    src="/references/slide-01.png"
    alt=""
  />
)}
```

Полезно добавить переключатель:

```js
document.addEventListener("keydown", (event) => {
  if (event.key.toLowerCase() === "d") {
    document.body.classList.toggle("debug-mode");
  }
});
```

Можно использовать три режима:

```text
1. Только HTML
2. Только оригинальный слайд
3. Наложение с прозрачностью 50%
```

Так быстро обнаруживаются:

* неверные координаты;
* несовпадающие размеры;
* ошибки шрифтов;
* неправильные интервалы;
* отличия в обрезке изображений.

---

# 17. Добавить навигацию между слайдами

Для лонгрида полезен вертикальный скролл со «прилипанием»:

```css
html {
  scroll-behavior: smooth;
}

.presentation {
  scroll-snap-type: y proximity;
}

.slide {
  scroll-snap-align: start;
}
```

Дополнительно можно сделать:

* точки навигации;
* номер текущего слайда;
* меню разделов;
* кнопки «вперед» и «назад»;
* управление стрелками клавиатуры;
* ссылку на конкретный слайд.

Пример URL:

```text
/report#slide-07
```

Каждому слайду присваивается ID:

```jsx
<section id={`slide-${index + 1}`}>
```

---

# 18. Сделать адаптацию под мобильные устройства

Здесь есть два подхода.

## Подход 1. Масштабировать слайд целиком

Самый точный способ.

Плюсы:

* дизайн полностью сохраняется;
* быстро реализуется;
* элементы не смещаются.

Минусы:

* на телефоне текст может быть мелким;
* пользователю придется увеличивать страницу.

## Подход 2. Создать мобильную версию каждого шаблона

Например:

```css
@media (max-width: 768px) {
  .slide-text-image__image {
    position: relative;
    width: 100%;
    height: 320px;
  }

  .slide-text-image__content {
    position: relative;
    width: 100%;
  }
}
```

Плюсы:

* удобно читать;
* хороший пользовательский опыт;
* можно превратить презентацию в настоящий лонгрид.

Минусы:

* мобильная версия уже не будет точной копией PPTX;
* потребуется дополнительный дизайн;
* объем разработки почти удвоится.

Практичный компромисс:

* на десктопе сохранять точный вид презентации;
* на телефоне переключаться на отдельную читабельную компоновку;
* содержимое брать из того же JSON.

---

# 19. Создать систему проверки данных

Чтобы шаблон не ломался, данные нужно валидировать.

Например, с помощью Zod:

```bash
npm install zod
```

```js
import { z } from "zod";

const CoverSlideSchema = z.object({
  id: z.string(),
  type: z.literal("cover"),
  title: z.string().min(1).max(100),
  subtitle: z.string().max(250).optional(),
  image: z.string().min(1)
});
```

Это поможет обнаружить:

* отсутствующее изображение;
* пустой заголовок;
* неправильный тип слайда;
* слишком длинный текст;
* недостаточное или лишнее количество карточек.

---

# 20. Продумать способ наполнения шаблона

Есть четыре возможных уровня.

## Самый простой

Редактировать JSON вручную:

```text
presentation.json
```

## Средний

Подключить Google Sheets или Excel, преобразовывая строки в JSON.

Например:

```text
slide_id | type | title | text | image
```

## Удобный

Сделать административную форму:

```text
Заголовок: [________________]
Текст:     [________________]
Изображение: [Загрузить]
```

## Продвинутый

Подключить CMS:

* Strapi;
* Directus;
* Sanity;
* Contentful;
* WordPress;
* собственную CMS.

Для первой версии лучше использовать JSON. Когда шаблон стабилизируется, поверх него можно построить редактор.

---

# 21. Предусмотреть экспорт обратно в PDF

HTML-лонгрид можно печатать в PDF.

CSS:

```css
@media print {
  @page {
    size: 13.333in 7.5in;
    margin: 0;
  }

  .slide {
    width: 13.333in;
    height: 7.5in;
    break-after: page;
    overflow: hidden;
  }

  .slide__canvas {
    transform: none;
  }

  .navigation,
  .debug-panel {
    display: none;
  }
}
```

Для автоматического экспорта можно использовать Playwright:

```bash
npm install playwright
```

```js
import { chromium } from "playwright";

const browser = await chromium.launch();
const page = await browser.newPage({
  viewport: {
    width: 1920,
    height: 1080
  }
});

await page.goto("http://localhost:5173", {
  waitUntil: "networkidle"
});

await page.pdf({
  path: "presentation.pdf",
  printBackground: true,
  width: "13.333in",
  height: "7.5in",
  margin: {
    top: "0",
    right: "0",
    bottom: "0",
    left: "0"
  }
});

await browser.close();
```

Это даст PDF, в котором каждая HTML-секция будет отдельной страницей.

---

# 22. Рекомендуемая структура проекта

```text
pptx-longread/
├── public/
│   ├── assets/
│   │   ├── images/
│   │   ├── icons/
│   │   └── fonts/
│   └── references/
│       ├── slide-01.png
│       ├── slide-02.png
│       └── slide-03.png
│
├── src/
│   ├── components/
│   │   ├── presentation/
│   │   │   ├── Presentation.jsx
│   │   │   ├── SlideFrame.jsx
│   │   │   └── SlideRenderer.jsx
│   │   │
│   │   └── slides/
│   │       ├── CoverSlide.jsx
│   │       ├── SectionSlide.jsx
│   │       ├── TextImageSlide.jsx
│   │       ├── CardsSlide.jsx
│   │       ├── StatisticsSlide.jsx
│   │       └── FinalSlide.jsx
│   │
│   ├── data/
│   │   └── presentation.json
│   │
│   ├── styles/
│   │   ├── reset.css
│   │   ├── tokens.css
│   │   ├── global.css
│   │   ├── slides.css
│   │   └── print.css
│   │
│   ├── utils/
│   │   ├── fitText.js
│   │   ├── scaleSlides.js
│   │   └── validateData.js
│   │
│   ├── App.jsx
│   └── main.jsx
│
├── scripts/
│   ├── inspect-pptx.py
│   ├── extract-media.py
│   └── export-pdf.js
│
├── package.json
└── vite.config.js
```

---

# 23. Последовательность работы над конкретной презентацией

## Этап 1. Аудит

1. Получить PPTX.
2. Определить размер презентации.
3. Экспортировать каждый слайд в PNG.
4. Извлечь изображения.
5. Собрать список шрифтов.
6. Выписать палитру.
7. Разбить слайды на типы.
8. Найти повторяющиеся элементы.

Результат этапа:

```text
8 типов слайдов
4 основных цвета
2 шрифта
26 исходных слайдов
18 изображений
```

## Этап 2. Техническая основа

1. Создать проект Vite + React.
2. Подключить CSS reset.
3. Создать дизайн-токены.
4. Создать `SlideFrame`.
5. Реализовать масштабирование.
6. Добавить вертикальный скролл.
7. Добавить отладочное наложение.

## Этап 3. Первый эталонный слайд

Начать с относительно простого слайда:

1. Разместить фон.
2. Подключить правильный шрифт.
3. Выставить координаты.
4. Настроить размеры текста.
5. Настроить изображение.
6. Сравнить с PNG.
7. Довести до близкого совпадения.

После первого слайда станет понятно, насколько правильно устроена система координат.

## Этап 4. Компоненты

1. Создать компонент обложки.
2. Создать текстовый слайд.
3. Создать текст + изображение.
4. Создать карточки.
5. Создать статистику.
6. Создать цитату.
7. Создать финальный слайд.

## Этап 5. Шаблонизация

1. Вынести текст из компонентов в JSON.
2. Вынести изображения в JSON.
3. Добавить варианты компонентов.
4. Добавить ограничения длины.
5. Добавить автоматическое уменьшение текста.
6. Добавить схему проверки данных.

## Этап 6. Контроль качества

Для каждого слайда проверить:

* фон;
* шрифты;
* размеры;
* межстрочные интервалы;
* переносы;
* координаты;
* изображения;
* скругления;
* тени;
* прозрачность;
* SVG;
* масштабирование;
* экспорт в PDF.

---

# 24. Как мы можем разделить работу

## Что делаете вы

1. Передаете PPTX.
2. Сообщаете, какие шрифты можно использовать.
3. Указываете, нужен ли мобильный вариант.
4. Определяете, какие поля должны быть переменными.
5. Уточняете, как будет заполняться шаблон:

   * JSON;
   * Excel;
   * форма;
   * CMS.
6. Проверяете визуальное совпадение.
7. Сообщаете, какие отклонения критичны.

## Чем помогаю я

Я могу последовательно:

1. Проанализировать структуру PPTX.
2. Определить типы слайдов.
3. Составить структуру данных.
4. Подготовить архитектуру проекта.
5. Написать React-компоненты.
6. Создать CSS для каждого типа слайда.
7. Сделать масштабирование.
8. Реализовать переменные.
9. Добавить подгонку текста.
10. Создать экспорт в PDF.
11. Проверить код.
12. Помочь довести HTML до визуального соответствия исходнику.

На практике удобнее работать по одному типу слайда:

```text
Сначала обложка → затем текстовый шаблон →
затем карточки → затем остальные варианты.
```

После создания одного компонента он применяется ко всем аналогичным слайдам.

---

# 25. Какой результат считать «максимально похожим»

Полностью идентичного отображения PowerPoint и браузера добиться не всегда возможно из-за различий в:

* рендеринге шрифтов;
* алгоритмах переноса строк;
* тенях;
* размытии;
* SVG;
* диаграммах;
* обработке прозрачности;
* специфических эффектах PowerPoint.

Реалистичные уровни:

### 90–95% совпадения

Достигается обычными HTML, CSS и SVG.

### 95–98% совпадения

Потребуются:

* точные шрифты;
* ручная настройка каждого типа слайда;
* наложение эталонных изображений;
* точная работа с координатами;
* отдельная настройка нестандартных форм.

### Практически 100%

Можно просто использовать слайды как фоновые изображения:

```css
.slide {
  background-image: url("/slides/slide-01.png");
}
```

Но тогда содержимое не будет нормально редактироваться и шаблонизироваться. Компромиссный вариант — оставить сложный декор фоном, а изменяемые тексты и изображения разместить поверх HTML-элементами.

---

# Рекомендуемый итоговый подход

Для вашей задачи я бы выбрал следующую архитектуру:

```text
Vite
React
Обычный CSS
JSON с данными
SVG для сложной графики
1920 × 1080 как внутренняя система координат
Пропорциональное масштабирование слайдов
Отдельные компоненты для 5–10 типов слайдов
Playwright для экспорта в PDF
```

Главный принцип:

```text
PPTX
  ↓
Анализ дизайна
  ↓
Набор типов слайдов
  ↓
React-компоненты
  ↓
Данные в JSON
  ↓
HTML-лонгрид
  ↓
При необходимости PDF
```

Первым практическим шагом стоит взять **один характерный слайд**, экспортировать его в PNG и создать точную HTML-копию. После этого можно утвердить систему координат, шрифты, способ масштабирования и структуру переменных — и уже на этой основе переносить всю презентацию.

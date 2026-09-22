# لعبة الثعبان 🐍 - Snake Game App

تطبيق أندرويد لجنلة الثعبان بأسلوب عصري ورسوم متحركة.

## خطوات تشغيل المشروع (بناء APK)

المشروع كامل وجاهز. تحتاج فقط:

### 1. تحميل الأدوات (مرة واحدة)
- **Android Studio** من: https://developer.android.com/studio
  - البرنامج يشتمل على Java و Android SDK تلقائيًا (يبنت معاه).

### 2. فتح المشروع
1. افتح **Android Studio**
2. File → Open → اختر المجلد `SnakeGameApp`
3. انتظر حتى ينتزل الأدوات (Gradle Sync) — أول مرة بتاخد وقت

### 3. بناء ملف التطبيق (APK)
1. من القائمة: **Build → Build Bundle(s) / APK(s) → Build APK(s)**
2. الملف هيظهر هنا:
   `SnakeGameApp/app/build/outputs/apk/debug/app-debug.apk`

### 4. تثبيت على هاتفك (للتجربة)
- انقل ملف الـ APK لهاتفك وافتحه للتثبيت
- أو استخدم USB:
  - فعّل Developer Mode على الهاتف
  - وصّله بالكمبيوتر
  - من Android Studio اضغط ▶ (Run)

---

## النشر على Google Play (لجني الأرباح 💰)

### أ) إنشاء حساب مطور
1. ادخل: https://play.google.com/console
2. سجّل بحساب Google خاص بك
3. ادفع رسوم التسجيل **$25** (مرة واحدة فقط)

### ب) تجهيز ملف النشر (AAB)
1. في Android Studio: **Build → Generate Signed Bundle / APK**
2. اختر **Android App Bundle (AAB)**
3. اعمل keystore جديد (كلمة مرور تحفظها لنفسك)
4. هيّصلك ملف `.aab`

### ج) رفع التطبيق
1. في Google Play Console → **Create app**
2. اعبي البيانات:
   - **الاسم**: لعبة الثعبان 🐍
   - **الوصف**: لعبة الثعبان الكلاسيكية بأسلوب عصري
   - **الفئة**: Games → Arcade
   - ارفع الصور (أيقونة + لقطات شاشة)
3. ارفع ملف الـ `.aab`
4. **Monetize**: اختر طريقة الربح (إعلانات أو شراء داخل التطبيق)
5. Submit for review (المراجعة بتاخد من يوم لأسبوع)

### ملاحظة مهمة
الاسم "باسمي" - أي بيانات النشر والحساب البنكي بتتسجل بالحساب نفسه. أنت اللي بتتحكم في نقل الأرباح لبنكك.

---

## الملفات الموجودة
```
SnakeGameApp/
├── app/
│   ├── src/main/
│   │   ├── AndroidManifest.xml       (إعدادات التطبيق)
│   │   ├── java/com/snakegame/app/   (الكود)
│   │   │   └── MainActivity.java
│   │   ├── assets/snake_game.html    (اللعبة نفسها)
│   │   └── res/                      (الرسومات والتنسيقات)
│   └── build.gradle
└── build.gradle
```

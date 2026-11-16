# Canvas Course Manager

Canvas Course Manager takes the lesson folders on your computer and builds your Canvas modules for you. It uploads lesson pages, quizzes, lesson plans, and answer keys, then names, orders, and indents them consistently. It can also update due dates, points, and module settings for a whole course from one spreadsheet.

You don't need any programming experience. Everything happens in a regular window with buttons and boxes. This guide walks you through setup one step at a time. Plan on about 20 minutes the first time.

**You can't break your course by trying it.** The program always shows you a preview first, never deletes anything in Canvas, and keeps answer keys hidden from students.

---

## What you need

- A Windows or Mac computer. Chromebooks and iPads won't work.
- A Canvas course where you are the teacher. A sandbox course is the perfect place to try it first.
- Your course files saved in folders on your computer. Step 5 shows how to arrange them.

---

## Step 1: Install Python (one time only)

Python is a free program that Canvas Course Manager runs on. You only install it once.

1. Go to **[python.org/downloads](https://www.python.org/downloads/)** and click the big yellow **Download Python** button.
2. Open the file you downloaded.

**On Windows:** On the first installer screen, check the box that says **Add python.exe to PATH** before clicking **Install Now**. This box is easy to miss, and the program won't open without it.

**On a Mac:** Click through the installer. When it finishes, a Finder window opens. Double-click **Install Certificates.command** in that window. This lets Python connect securely to Canvas.

Already have Python? It needs to be version 3.10 or newer. If yours is too old, the program will tell you.

---

## Step 2: Download Canvas Course Manager

1. Open the download page: **https://github.com/odysseyCTE-hub/Canvas-Course-Manager/releases/latest**
2. Under **Assets**, click **Source code (zip)**.
3. Unzip the downloaded file. On Windows, right-click it and choose **Extract All**. On a Mac, double-click it.
4. Move the unzipped folder somewhere easy to find, such as your **Documents** folder.

The program can live in any folder, and one copy can manage all your courses.

---

## Step 3: Open the program

**On Windows:** Double-click **Start Canvas Course Manager.bat**. If a blue "Windows protected your PC" box appears, click **More info**, then **Run anyway**.

**On a Mac:** Double-click **Start Canvas Course Manager.command**. The first time, your Mac may say it can't verify the file. If so, right-click the file, choose **Open**, and then click **Open** again. A Terminal window opens alongside the program. Leave it open while you work.

**The very first time, setup happens automatically.** A plain text window appears and installs a few free helper pieces that run the settings spreadsheet and give the program its modern look. Make sure you're connected to the internet, and don't close that window. It takes a minute or two and only happens once. If the program ever asks to install anything, click **Yes**. You need these pieces.

---

## Step 4: Connect to Canvas (the Course setup tab)

The program needs three things from Canvas: your school's Canvas address, an access token, and your course ID.

### Your Canvas address

This is the web address you see when you log in to Canvas, such as `https://yourschool.instructure.com`. Copy just that first part.

### An access token

A token is like a special password that lets the program work in Canvas as you.

1. In Canvas, click **Account** (your profile picture on the left), then **Settings**.
2. Scroll to **Approved Integrations** and click **+ New Access Token**.
3. For **Purpose**, type `Canvas Course Manager`. You can leave the expiration date blank, or pick one if your school prefers.
4. Click **Generate Token**.
5. **Copy the long token right away.** Canvas only shows it once. If you lose it, just make a new one.

Paste the address and the token into the **Course setup** tab. The program saves them in a file called `canvas_config.txt` next to the program.

> 🔒 **Keep your token private.** Anyone with it can act as you in Canvas. Never email, upload, or share `canvas_config.txt`. If you think it got out, go back to Canvas Settings, delete that token, and make a new one.

### Your course ID

Open your course in Canvas and look at the web address. The number after `/courses/` is your course ID. For example, in `https://yourschool.instructure.com/courses/1234`, the course ID is **1234**.

### Your course folder

Click the button to choose the folder that holds this course's files. The program remembers your recent courses, so next time you can pick from a list. Each course folder keeps its own settings.

---

## Step 5: Organize your course folder

The program finds your files by folder and file names. Here's an example layout:

```
Fundamentals of Computing/          ← the course folder you choose in the program
├── Unit 1 - Computer Basics/
│   ├── Lesson 1/
│   │   ├── Lesson 1 Plan.docx               → lesson plan (goes to the "Lesson Plans" module)
│   │   ├── Parts of a Computer.html          → lesson page
│   │   ├── Lesson 1 Slides.pptx              → presentation
│   │   ├── Lesson 1 Checkpoint.zip           → quiz (becomes a Canvas New Quiz)
│   │   └── Lesson 1 Checkpoint Answer Key.pdf → answer key (always hidden from students)
│   ├── Lesson 2/
│   │   └── ...
│   ├── Unit 1 Activity/
│   │   ├── Unit 1 Activity.zip
│   │   └── Unit 1 Activity Answer Key.pdf
│   └── Unit 1 Exam/
│       ├── Unit 1 Exam.zip
│       └── Unit 1 Exam Answer Key.pdf
└── Unit 2 - Networks/
    └── ...
```

**The naming rules:**

- **Unit folders** start with the word *Unit* and a number. You can add a name after it, as in "Unit 1 - Computer Basics." Each unit folder matches the Canvas module with the same number.
- **Lesson folders** go inside unit folders and start with *Lesson* and a number.
- **Lesson plans** are Word files (`.doc` or `.docx`) with *plan* in the name.
- **Lesson pages** are `.html` files.
- **Quizzes** are QTI `.zip` files, the kind exported from a quiz-making tool. Their names include *Checkpoint*, *Activity*, or *Exam*, or they sit in an Activity or Exam folder. Other zip files are ignored.
- **Answer keys** are PDFs with *Answer Key* in the name. They are always locked so students can't see them.
- **Presentations** are PowerPoint files in a lesson folder.

**Do your courses use different words?** If you say "Module" or "Chapter" instead of "Unit," or "Day" or "Week" instead of "Lesson," change them on the **Names & look** tab. That tab also lets you change or turn off the emojis in item titles (📄 📚 📝 🪛 ⭐) and adjust indentation. A live sample module shows exactly how your choices will look before anything changes in Canvas.

---

## Step 6: Preview, then apply (the Run tab)

1. Choose **what you'd like to do** and **where the program should look**: one lesson, one unit, or the whole course.
2. Leave **Preview only (recommended)** selected and click **Run preview**. The program shows what it *would* do and changes nothing in Canvas.
3. Read the preview. If it looks right, select **Apply changes to Canvas**, click **Apply changes…**, and type the word `APPLY` when asked.

Run again any time you edit files. The program remembers what it already uploaded and skips anything that hasn't changed.

**Tip:** New modules and pages start out unpublished, so students won't see anything until you publish it.

---

## The settings spreadsheet (optional)

The spreadsheet lets you update many things at once: due dates, points, assignment groups, item titles and order, what's published, completion requirements, and prerequisites.

- **Don't build it by hand.** Choose **Create settings spreadsheet** on the Run tab. The program reads your course from Canvas (it only looks, it doesn't change anything) and makes a filled-in workbook in your course folder. You can also choose a blank template.
- Every row starts as **Apply changes = NO**. Change a row to **YES** only when you want that change made.
- The program never overwrites an existing spreadsheet. Each new one is saved as a separate, dated file.
- As always, run a preview first.

## Assignment groups for quizzes

You can choose which Canvas assignment group your Checkpoint, Activity, and Exam quizzes go into, or let Canvas decide. The program loads your existing groups from Canvas and never creates a new group unless you click to make one yourself. It shows you a preview first, and it warns you if a new group would count for 0% of the grade.

---

## Updates

When a new version is ready, the program will offer to install it when you open it. Click **Update now** to accept, or skip it until later.

Updates only replace the program itself. Your token, your settings, and your course files are never touched. Your old copy is saved in the `backups` folder, just in case.

---

## The safety rules

These are built in and can't be turned off:

- **Preview first.** Nothing changes in Canvas until you choose Apply and type `APPLY`.
- **Nothing is ever deleted** in Canvas.
- **Answer keys are always hidden** from students.
- **No guessing.** If a title could match more than one item in Canvas, the program skips it and tells you, instead of picking one.
- **Everything is recorded.** Each course folder has a log file (`canvas_lesson_uploader.log`) listing every change. Click **Open full log** on the Run tab to see it.

---

## Something not working?

| What you see | What to try |
|---|---|
| The launcher says it can't find Python (Windows) | Python may not be installed, or the "Add python.exe to PATH" box wasn't checked. Run the Python installer again, choose **Modify**, and make sure that option is on. |
| The Mac says the launcher can't be opened or I don't have permission | Right-click **canvas_course_manager.py** and choose **Open With → Python Launcher**. Then let Lisa know. |
| The setup window says the helper packages didn't install | Check your internet connection and open the launcher again. School networks sometimes block installs; if it keeps failing, try from home. |
| "Needs Python 3.10 or newer" | Install the newest Python from python.org (Step 1). |
| "Could not open Canvas course" | Check the course ID on the Course setup tab, and make sure you're a teacher in that course. |
| Canvas says my login doesn't work | Your token may have expired or been deleted. Make a new one (Step 4) and paste it in. |
| A lesson or file didn't show up in the preview | Check the folder and file names against the rules in Step 5, or the words on the Names & look tab. |
| Something was skipped as "unclear" | Two Canvas items have similar titles. Rename one in Canvas so they're easier to tell apart, then run again. |

Still stuck? Ask Lisa, and send along the log file from your course folder. It shows exactly what happened. (Never send `canvas_config.txt`.)

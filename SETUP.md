# Tabeeby website: setup guide

## What's in the folder

| File | What it does |
|---|---|
| `index.html` | The whole one-page website. Arabic is the default language and English is one click away. |
| `content.xlsx` | **The content file** (or its Google Sheet copy, see section 1). Departments, treatments, doctors, FAQ, phones, WhatsApp, address, texts, colours and booking links all live here. |
| `content-fallback.js` | A backup copy of the Excel content. The site only uses it when it is opened straight from your computer. |
| `logo.png`, `img/` | Logo, white logo, favicon, doctor photos (`img/doctors/`) and department photos (`img/departments/`). |
| `src/`, `tools/` | Developer files. You don't need them to run the site. |

To put the site online, upload everything (except `src/` and `tools/` if you like) to any static host, such as Vercel, Netlify, GitHub Pages or cPanel hosting.

---

## 1. Edit from Google Sheets (no uploading)

Do this once, then every change you make in the sheet appears on the website by itself.

1. Go to [sheets.google.com](https://sheets.google.com) and open a blank sheet, then go to **File → Import → Upload** and choose `content.xlsx`. Pick **Replace spreadsheet**.
2. Go to **File → Share → Publish to web**. Choose **Entire document** and **Microsoft Excel (.xlsx)**, then click **Publish** and copy the link.
3. Open `index.html` in any text editor. Near the top, paste the link between the quotes:
   `GOOGLE_SHEET_URL: "paste-the-link-here",`
4. Upload `index.html` to the website one last time.

From now on you only edit the Google Sheet. Google refreshes the published copy every few minutes, so changes show up on the site within about 5 minutes.

- Keep the tab names and the column titles in row 1 exactly as they are.
- Publishing makes the sheet readable by anyone with the link. That's fine here because everything in it is already shown on the website, so don't add private notes to it.
- If Google can't be reached, the site automatically uses `content.xlsx` from the website folder instead. Uploading a fresh copy now and then (**File → Download → Microsoft Excel**) keeps that backup current.

---

## 2. Managing the content (Excel or Google Sheet)

1. Open `content.xlsx` in Excel or Google Sheets.
2. Edit the tab you need:
   - **Settings**: phone, WhatsApp, email, address, hours, hero texts, the About text, colours and the main booking link.
   - **Departments**: one row per department. Write `no` in `show` to hide one.
   - **Services**: the treatments shown inside each department, such as kids' dentistry, veneers, crowns, whitening, dentures and skin treatments.
   - **Doctors**: one row per doctor. The `department` column must match a department `id`.
   - **Features**, **FAQ**: the short benefit cards and the questions.
3. Save it and upload it to the website folder, replacing the old `content.xlsx`. Keep the same file name.

Don't rename the tabs or the column titles in row 1. Lower `order` numbers appear first.

**Doctor photos:** put the photo in `img/doctors/`, for example `img/doctors/nadia.jpg`, then write that path in the doctor's `photo` cell. A full `https://` image link also works. Until a photo is added, the site shows the doctor's initials.

---

## 3. Google Calendar bookings

Patients see your real free times and book them. Each booking lands in your Google Calendar with the patient's name, phone and email, and Google sends them a confirmation email with a link to cancel or reschedule. You manage everything from Google Calendar only.

### Create a booking page

1. Open **Google Calendar** on a computer, signed in with the clinic's Google account.
2. Click **Create** and choose **Appointment schedule**.
3. Set the title (for example *Dr. Nadia Mohamed – Dermatology*), the appointment length and your weekly available hours.
4. Open **Booked appointment settings** and require the **phone number** field. You can also add a question such as *Patient age* or *Treatment*.
5. Under **Scheduling window** and **Buffer time**, set how far ahead people may book and the gap between visits.
6. Click **Save**.

### Get the link for the website

1. Click the schedule in your calendar and choose **Share**.
2. Choose **Website embed → Inline schedule** and copy the address inside `src="…"`. It looks like this:
   `https://calendar.google.com/calendar/appointments/schedules/AcZssZ…?gv=true`
3. Paste it into the sheet (Google Sheet or `content.xlsx`):
   - One shared calendar for the whole clinic: **Settings → BOOKING_URL**.
   - One calendar per department: the department's `booking_url`.
   - One calendar per doctor (recommended): the doctor's `booking_url`.

The site uses the doctor's link first, then the department's, then `BOOKING_URL`. Patients who choose "Any available doctor" see the department's calendar.

Until a booking link is added, the booking area shows a short request form that opens WhatsApp with the patient's details already typed in, so no request is lost while you set up the calendar.

**Good to know**

- Use the full `calendar.google.com/...` link from *Website embed*. Short `calendar.app.google/...` links can't be shown inside the page.
- A free Google account can create appointment schedules. Some extras, such as multiple booking pages with payment or automatic reminder emails, need Google Workspace or Google One Premium.
- Turn on Google Calendar notifications on your phone so you get an alert for every new booking.

---

## 4. WhatsApp button

The green floating button and every WhatsApp link use **Settings → WHATSAPP**. Write digits only, with the country code: `971503909206`.

## 5. Language

The site opens in Arabic. Visitors switch with the EN / عربي button, and their choice is remembered. Direct links: `?lang=ar` or `?lang=en`. Links to a doctor's booking: `?doctor=nadia-mohamed#book`.

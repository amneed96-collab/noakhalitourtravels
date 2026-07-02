// ============================================================
// কনফিগারেশন - এখানে আপনার Google Apps Script Web App URL বসান
// ============================================================
//
// কীভাবে পাবেন:
// 1. Google Sheet খুলুন > Extensions > Apps Script
// 2. Code.gs এ ব্যাকএন্ড কোড পেস্ট করুন
// 3. setupSheets ফাংশন একবার রান করুন
// 4. Deploy > New deployment > Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 5. যে URL পাবেন (.../exec দিয়ে শেষ) সেটি নিচে বসান

const API_URL = 'https://script.google.com/macros/s/AKfycbzai1VLFxw9PpcGRfsRbFu5EkFsLrjmyTLi0k0dt4YMHpdjgJV2V1xwKHMYi-uPvL5e/exec';

// গ্রুপের নাম ও তথ্য (চাইলে পরিবর্তন করুন)
const SITE_CONFIG = {
  groupName: 'নোয়াখালী ট্যুর ট্রাভেলস',
  tagline: 'বাংলাদেশ ঘুরে দেখি, একসাথে',
  facebookUrl: 'https://web.facebook.com/abdul.momen.111408',
  whatsappNumber: '01605721296',
  phone: '01605721296',
};

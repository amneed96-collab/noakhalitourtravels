// ============================================================
// কনফিগারেশন - এখানে আপনার Google Apps Script Web App URL বসান
// ============================================================
//
// 1. Google Sheet > Extensions > Apps Script এ Code.gs পেস্ট করুন
// 2. setupSheets ফাংশন একবার রান করুন
// 3. Deploy > New deployment > Web app (Execute as: Me, Access: Anyone)
// 4. যে URL পাবেন সেটি নিচে বসান

const API_URL = 'https://script.google.com/macros/s/AKfycbzai1VLFxw9PpcGRfsRbFu5EkFsLrjmyTLi0k0dt4YMHpdjgJV2V1xwKHMYi-uPvL5e/exec';

const SITE_CONFIG = {
  groupNameBn: 'ট্রাভেল গ্রুপ অফ নোয়াখালী',
  groupNameEn: 'Travel Group of Noakhali',
  tagline: 'বাংলাদেশ ঘুরে দেখি, একসাথে',

  // সোশ্যাল মিডিয়া লিংক (ফুটারে বাম পাশে দেখাবে)
  facebookUrl: 'https://web.facebook.com/abdul.momen.111408',
  youtubeUrl: '',
  instagramUrl: '',

  whatsappNumber: '01794709881',
  phone: '01853302933',

  // লোগো - ছবির লিংক বসান (ফাঁকা রাখলে ডিফল্ট 🧳 আইকন দেখাবে)
logoUrl: 'https://drive.google.com/thumbnail?id=17-XQUHTUc_fvNQi16iKfVGg0xDFHmXMg&sz=w1000',

  // বিকাশ/নগদ নম্বর (বুকিং ফরমে দেখানো হবে)
  bkashNumber: '01794709881',
  nagadNumber: '01794709881',

  // হেডারে দেখানো তথ্য
  headerAdminLine: 'এডমিন: মোহাম্মদ ওয়াসিম- 01794709881 / তাসিব- 01853302933',

  // ফুটারে সবার নিচে ও প্রতিটি টিকেট/প্রিন্ট ফরমের নিচে দেখানো ক্রেডিট লাইন
  poweredByLine: '@PowredBy: AMShahed-01605721296'
};

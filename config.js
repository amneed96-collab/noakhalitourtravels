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
  groupName: 'নোয়াখালী ট্যুর ট্রাভেল্স',
  tagline: 'বাংলাদেশ ঘুরে দেখি, একসাথে',
  facebookUrl: 'https://web.facebook.com/abdul.momen.111408',
  whatsappNumber: '01605721296',
  phone: '01605721296',

  // হেডারে দেখানো এডমিন তথ্য
  headerAdminLine: 'এডমিন- রবিউল হক ওয়াসিম: 01648-093805',

  // ফুটারে ও টিকেটে দেখানো তথ্য
  footerLine: '@AMShahed: 01605721296',
  ticketAdminName: 'রবিউল হক ওয়াসিম',
  ticketAdminPhone: '01648-093805',
  ticketFooterCredit: '@AMShahed: 01605721296'
};

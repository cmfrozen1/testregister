// ==============================================
// Code.gs - ระบบลงทะเบียน + LINE Webhook + Member Card (Reply API — ฟรี ไม่เสียโควต้า)
// ==============================================

const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
const LINE_CHANNEL_ACCESS_TOKEN = "oVCQJNRQYKH6Dfe75lmrygNDkVVnaqDoUyu4crkUPfMpc88Vhm59VXbOEz5lQmO2y4rYvYf326/Gno9djAJlYDAydHocMRpsfs/9lka8PZQyPJCdJi3ohXU3qYxIHD5tJdaAPh1FxkyrcXCoaE0ptwdB04t89/1O/w1cDnyilFU="; // 👈 AccessToken ของบอท LINE OA (Map_API)
const REGISTER_FORM_URL = "https://testregister-ten.vercel.app/";
const WELCOME_POINTS = 10;           // 👈 แต้มต้อนรับตอนลงทะเบียนสำเร็จ ปรับได้ตรงนี้
const MEMBER_VALID_YEARS = 1;        // 👈 อายุสมาชิก (ปี) ก่อนหมดอายุ
const SEND_F_AFTER_REGISTER = true;  // 👈 ส่งข้อความ "F" หลังลงทะเบียนสำเร็จ (meme) — ปิดได้ตรงนี้
const F_MESSAGE_TEXT = "F";         // 👈 ข้อความที่จะส่ง (เช่น "F" หรือ "F 🙏")
const REPLY_INVITE_ON_FOLLOW = false; // 👈 หลักการคลิป: false = เก็บ replyToken ไว้ตอบบัตรหลังลงทะเบียน (ฟรี ไม่ใช้ push)
                                      //     true  = ตอบคำเชิญเองตอนแอดเพื่อน (จะกิน replyToken — ไม่แนะนำ)

// ==============================================
// 🔗 LINE Login (Bot Link) — เชื่อมต่อกับ LINE ทางการ
// หน้าจอ "เชื่อมต่อกับ LINE" เป็นบริการทางการของ LINE (LINE Login)
// หลังเชื่อมต่อ LINE จะส่ง link event (มี replyToken) → บอท Reply บัตรสมาชิก (ฟรี)
// ==============================================
const LINE_LOGIN_CHANNEL_ID = "2000688983";                    // 👈 LINE Developers → LINE Login channel → Channel ID
const LINE_LOGIN_CHANNEL_SECRET = "8d26d65c8c4152fef14d060541140cd3"; // 👈 LINE Login channel → Channel Secret
const FORM_REDIRECT_URI = "https://testregister-ten.vercel.app/";   // 👈 ต้องตรงกับ redirect_uri ที่ลงทะเบียนใน LINE Login channel

// ==============================================
// doPost - รับข้อมูลทั้งจาก LINE Webhook และฟอร์ม HTML
// ==============================================

function doPost(e) {
  try {
    // ✅ บันทึกข้อมูลดิบเพื่อ Debug
    const props = PropertiesService.getScriptProperties();
    if (e && e.postData) {
      props.setProperty("last_raw_data", e.postData.contents);
    }

    Logger.log("📩 doPost ถูกเรียกใช้งาน");
    Logger.log("📩 e: " + JSON.stringify(e));

    const isLineWebhook = e && e.postData && e.postData.contents &&
      e.postData.contents.includes('"events"') &&
      e.postData.contents.includes('"replyToken"');

    Logger.log("🔍 isLineWebhook: " + isLineWebhook);

    if (isLineWebhook) {
      return handleLineWebhook(e);
    } else {
      return handleFormSubmit(e);
    }
  } catch (error) {
    Logger.log("❌ doPost Error: " + error.message);
    Logger.log("❌ Stack: " + error.stack);
    return ContentService.createTextOutput("❌ เกิดข้อผิดพลาด: " + error.message);
  }
}

// ==============================================
// จัดการ LINE Webhook
// ==============================================

function handleLineWebhook(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const events = payload.events;

    Logger.log("📩 รับข้อมูลจาก LINE: " + JSON.stringify(payload));

    if (!events || events.length === 0) {
      Logger.log("⚠️ ไม่มี Events");
      return ContentService.createTextOutput("OK");
    }

    events.forEach(event => {
      try {
        Logger.log("📌 Event Type: " + event.type);

        // ✅ ผู้ใช้เพิ่มเพื่อน LINE OA (จุดเริ่มต้นของการเชื่อมต่อ)
        if (event.type === 'follow') {
          const replyToken = event.replyToken;
          const userId = event.source.userId;

          Logger.log("👋 ผู้ใช้เพิ่มเพื่อน (follow): " + userId);

          const props = PropertiesService.getScriptProperties();
          props.setProperty("token_" + userId, replyToken);
          props.setProperty("ts_" + userId, String(Date.now()));

          // ✅ หลักการจากคลิป: เก็บ replyToken ไว้ (ยังไม่ใช้!) เพื่อ Reply บัตรสมาชิก
          //    อัตโนมัติหลังลงทะเบียนภายใน ~1 นาที (Reply API ฟรี ไม่เสียโควต้า push)
          //    คำเชิญ/ลิงก์ฟอร์ม มาจากข้อความต้อนรับ (Welcome Message) ใน LINE OA Manager

          // ถ้าลงทะเบียนไว้ก่อนแล้ว (ผ่านลิงก์ ?uid=) → ตอบกลับด้วยบัตรสมาชิกเลย (ใช้ token นี้)
          const memberJson = props.getProperty("member_" + userId);
          if (memberJson) {
            replyMemberCardFlex(replyToken, JSON.parse(memberJson));
            props.deleteProperty("member_" + userId);
            props.deleteProperty("token_" + userId);
            props.deleteProperty("ts_" + userId);
          } else if (REPLY_INVITE_ON_FOLLOW) {
            // (ทางเลือก) ส่ง Flex เชิญลงทะเบียนทันที — จะกิน replyToken
            replyFlexMessage(replyToken, userId);
          } else {
            Logger.log("💾 เก็บ replyToken ไว้ตอบบัตรหลังลงทะเบียน (คำเชิญมาจาก Welcome Message ใน OA Manager)");
          }
          return;
        }

        // ✅ ผู้ใช้เชื่อมต่อ LINE ทางการผ่าน LINE Login (Bot Link) — เก็บ replyToken ไว้
        if (event.type === 'link') {
          const replyToken = event.replyToken;
          const userId = event.source.userId;
          Logger.log("🔗 ผู้ใช้เชื่อมต่อ LINE (link): " + userId);

          const props = PropertiesService.getScriptProperties();
          props.setProperty("token_" + userId, replyToken);
          props.setProperty("ts_" + userId, String(Date.now()));

          // ถ้ามีบัตรรอส่งอยู่ (ลงทะเบียนก่อน แล้วเพิ่งเชื่อมต่อ LINE) → Reply เลย (ฟรี)
          const pendingReply = props.getProperty("pending_reply_" + userId);
          if (pendingReply) {
            replyMemberCardFlex(replyToken, JSON.parse(pendingReply));
            props.deleteProperty("pending_reply_" + userId);
            props.deleteProperty("token_" + userId);
            props.deleteProperty("ts_" + userId);
            Logger.log("📤 ส่งบัตรสมาชิกที่รออยู่ผ่าน link event (ฟรี)");
          }
          return;
        }

        // ✅ ผู้ใช้บล็อก/ลบเพื่อน LINE OA (ยกเลิกการเชื่อมต่อ)
        if (event.type === 'unfollow') {
          const userId = event.source.userId;
          Logger.log("🚫 ผู้ใช้บล็อก/ลบเพื่อน (unfollow): " + userId);

          const props = PropertiesService.getScriptProperties();
          props.deleteProperty("token_" + userId);
          props.deleteProperty("ts_" + userId);
          props.deleteProperty("member_" + userId);
          return;
        }

        if (event.type === 'message' && event.message.type === 'text') {
          const replyToken = event.replyToken;
          const userId = event.source.userId;
          const userMessage = event.message.text;

          Logger.log("✅ replyToken: " + replyToken);
          Logger.log("👤 userId: " + userId);
          Logger.log("💬 ข้อความ: " + userMessage);

          // ✅ เก็บ replyToken
          const props = PropertiesService.getScriptProperties();
          props.setProperty("token_" + userId, replyToken);
          props.setProperty("ts_" + userId, String(Date.now()));
          props.setProperty("last_replyToken", replyToken);
          props.setProperty("last_userId", userId);
          props.setProperty("last_message", userMessage);

          // ✅ เทคนิคจากคลิป (0:26): ถ้าผู้ใช้ลงทะเบียนแล้ว → ตอบกลับด้วยบัตรสมาชิก
          //    (Reply API ฟรี ไม่เสียโควต้า push — ได้ replyToken ใหม่ทุกครั้งที่ผู้ใช้พิมพ์)
          const memberJson = props.getProperty("member_" + userId);
          if (memberJson) {
            replyMemberCardFlex(replyToken, JSON.parse(memberJson));
            props.deleteProperty("member_" + userId); // ส่งครั้งเดียว
          } else {
            // ✅ ส่ง Flex Message (replyToken ใช้ได้ครั้งเดียว
            //  จึงไม่ควรส่งข้อความทดสอบก่อน เพราะจะกิน token ไปก่อน)
            replyFlexMessage(replyToken, userId);
          }
        } else if (event.type !== 'follow' && event.type !== 'unfollow') {
          Logger.log("⚠️ Event ไม่ใช่ข้อความ: " + event.type);
        }
      } catch (err) {
        Logger.log("⚠️ Event ผิดพลาด (ไม่หยุดทั้ง batch): " + err.message);
      }
    });

    return ContentService.createTextOutput("OK");
  } catch (error) {
    Logger.log("❌ Webhook Error: " + error.message);
    Logger.log("❌ Stack: " + error.stack);
    return ContentService.createTextOutput("❌ Webhook Error: " + error.message);
  }
}

// ==============================================
// ส่ง Flex Message (Reply) - เชิญไปกรอกฟอร์ม
// ==============================================

function replyFlexMessage(replyToken, userId) {
  Logger.log("📤 กำลังส่ง Flex Message...");

  if (!LINE_CHANNEL_ACCESS_TOKEN || LINE_CHANNEL_ACCESS_TOKEN === "YOUR_LONG_LIVED_ACCESS_TOKEN") {
    Logger.log("⚠️ ยังไม่ได้ตั้งค่า LINE_CHANNEL_ACCESS_TOKEN");
    return;
  }

  const url = "https://api.line.me/v2/bot/message/reply";
  const formUrl = REGISTER_FORM_URL + "?uid=" + encodeURIComponent(userId);

  const flexMessage = {
    replyToken: replyToken,
    messages: [{
      type: "flex",
      altText: "✨ ยินดีต้อนรับสู่ LINE OA",
      contents: {
        type: "bubble",
        size: "mega",
        header: {
          type: "box",
          layout: "vertical",
          paddingAll: "20px",
          background: {
            type: "linearGradient",
            angle: "160deg",
            startColor: "#8b5cf6",
            endColor: "#4f46e5"
          },
          contents: [
            {
              type: "text",
              text: "✨ ยินดีต้อนรับ",
              color: "#ffffff",
              weight: "bold",
              size: "xl"
            },
            {
              type: "text",
              text: "คุณมาถูกที่แล้ว! มาเริ่มกันเลย",
              color: "#e0e7ff",
              size: "sm",
              margin: "xs"
            }
          ]
        },
        body: {
          type: "box",
          layout: "vertical",
          spacing: "lg",
          paddingAll: "20px",
          contents: [
            {
              type: "box",
              layout: "vertical",
              paddingAll: "16px",
              spacing: "xs",
              cornerRadius: "12px",
              background: {
                type: "linearGradient",
                angle: "90deg",
                startColor: "#fef3c7",
                endColor: "#ffedd5"
              },
              contents: [
                { type: "text", text: "🎁", size: "3xl", align: "center" },
                {
                  type: "text",
                  text: "รับบัตรสมาชิก + แต้มต้อนรับ " + WELCOME_POINTS + " คะแนน ฟรี!",
                  size: "sm",
                  weight: "bold",
                  color: "#78350f",
                  align: "center",
                  wrap: true
                }
              ]
            },
            {
              type: "box",
              layout: "horizontal",
              spacing: "md",
              contents: [
                { type: "text", text: "🎫", size: "md" },
                {
                  type: "text",
                  text: "บัตรสมาชิกดิจิทัล ใช้แทนบัตรจริงได้",
                  size: "sm",
                  color: "#374151",
                  gravity: "center",
                  flex: 1,
                  wrap: true
                }
              ]
            },
            {
              type: "box",
              layout: "horizontal",
              spacing: "md",
              contents: [
                { type: "text", text: "⭐", size: "md" },
                {
                  type: "text",
                  text: "สะสมคะแนนทุกครั้งที่ใช้บริการ",
                  size: "sm",
                  color: "#374151",
                  gravity: "center",
                  flex: 1,
                  wrap: true
                }
              ]
            },
            {
              type: "box",
              layout: "horizontal",
              spacing: "md",
              contents: [
                { type: "text", text: "🔔", size: "md" },
                {
                  type: "text",
                  text: "รับข่าวสารและโปรโมชันพิเศษก่อนใคร",
                  size: "sm",
                  color: "#374151",
                  gravity: "center",
                  flex: 1,
                  wrap: true
                }
              ]
            },
            {
              type: "button",
              style: "primary",
              height: "sm",
              color: "#4f46e5",
              action: {
                type: "uri",
                label: "📝 ไปที่หน้าลงทะเบียน",
                uri: formUrl
              }
            }
          ]
        },
        footer: {
          type: "box",
          layout: "vertical",
          paddingAll: "12px",
          contents: [
            {
              type: "text",
              text: "⏳ ใช้เวลาเพียง 1 นาที ลงทะเบียนครั้งเดียว รับสิทธิ์ทันที",
              size: "xs",
              color: "#9ca3af",
              align: "center",
              wrap: true
            }
          ]
        }
      }
    }]
  };

  const options = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify(flexMessage),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = response.getContentText();
    Logger.log("✅ ส่ง Flex Message สำเร็จ");
    Logger.log("📬 Response: " + result);
  } catch (error) {
    Logger.log("❌ ส่ง Flex Message ล้มเหลว: " + error.message);
  }
}

// ==============================================
// จัดการข้อมูลจากฟอร์ม HTML
// ==============================================

function handleFormSubmit(e) {
  try {
    const data = e.parameter;

    Logger.log("📩 ข้อมูลจากฟอร์ม: " + JSON.stringify(data));

    if (!data.firstname || !data.lastname || !data.email || !data.phone) {
      return ContentService.createTextOutput("❌ กรุณากรอกข้อมูลให้ครบถ้วน");
    }

    // บันทึกข้อมูลลง Google Sheet
    sheet.appendRow([
      new Date(),
      data.firstname,
      data.lastname,
      data.email,
      data.phone,
      data.gender || "ไม่ระบุ",
      data.interest || "",
      data.agree || "ไม่"
    ]);

    const userId = data.lineUserId || "";
    Logger.log("👤 lineUserId จากฟอร์ม: " + userId);

    let message = "✅ บันทึกข้อมูลสำเร็จ";

    // สร้าง state สำหรับเชื่อมต่อ LINE (LINE Login) — บัตรจะถูกส่งหลังเชื่อมต่อเสร็จ
    const state = Utilities.getUuid();
    const expireDate = new Date();
    expireDate.setFullYear(expireDate.getFullYear() + MEMBER_VALID_YEARS);

    const member = {
      name: data.firstname + " " + data.lastname,
      points: WELCOME_POINTS,
      phone: data.phone,
      expireDate: formatThaiDate(expireDate)
    };

    if (userId) {
      // ✅ มี userId แล้ว → ใช้ replyToken ที่เก็บไว้ตอนแอดเพื่อน/เชื่อมต่อ LINE
      //    Reply บัตรสมาชิกอัตโนมัติทันที (ฟรี ไม่เสียโควต้า push)
      const props = PropertiesService.getScriptProperties();

      // เก็บข้อมูลบัตรไว้ (fallback: ถ้า token หมดอายุ ผู้ใช้พิมพ์ข้อความในแชทเพื่อรับบัตร)
      props.setProperty("member_" + userId, JSON.stringify(member));
      Logger.log("💾 เก็บข้อมูลบัตรสมาชิกสำหรับ userId: " + userId);

      const savedToken = props.getProperty("token_" + userId);
      if (savedToken) {
        const sent = replyMemberCardFlex(savedToken, member);
        props.deleteProperty("token_" + userId);
        props.deleteProperty("ts_" + userId);
        if (sent) {
          props.deleteProperty("member_" + userId); // ส่งสำเร็จ → ไม่ต้องส่งซ้ำ
          message = "✅ บันทึกข้อมูลสำเร็จ — ส่งบัตรสมาชิกไปยัง LINE แล้ว (ฟรี ไม่เสียโควต้า)";
        } else {
          message = "✅ บันทึกข้อมูลสำเร็จ — replyToken หมดอายุแล้ว พิมพ์ข้อความในแชท LINE เพื่อรับบัตรสมาชิก";
        }
      } else {
        message = "✅ บันทึกข้อมูลสำเร็จ — พิมพ์ข้อความในแชท LINE เพื่อรับบัตรสมาชิก (ฟรี ไม่เสียโควต้า)";
      }
    } else {
      // ✅ ยังไม่เชื่อม LINE → เก็บข้อมูลรอไว้ แล้วเด้งไปหน้าจอ "เชื่อมต่อกับ LINE" ทางการ
      //    พอเชื่อมต่อเสร็จ LINE ส่ง link event → GAS Reply บัตรสมาชิกให้ (ฟรี)
      const props = PropertiesService.getScriptProperties();
      props.setProperty("pending_" + state, JSON.stringify(member));
      Logger.log("💾 เก็บข้อมูลรอเชื่อมต่อ LINE (state=" + state + ")");
      message = "✅ บันทึกข้อมูลสำเร็จ — กรุณาเชื่อมต่อ LINE เพื่อรับบัตรสมาชิก";
    }

    return ContentService.createTextOutput(message + "\nSTATE:" + state);

  } catch (error) {
    Logger.log("❌ handleFormSubmit Error: " + error.message);
    Logger.log("❌ Stack: " + error.stack);
    return ContentService.createTextOutput("❌ เกิดข้อผิดพลาด: " + error.message);
  }
}

// ==============================================
// Flex Message: บัตรสมาชิก + คะแนนสะสม (Reply API)
// เทคนิคจากคลิป (0:26): เปิดแชท LINE OA แล้วพิมพ์ข้อความ → ตอบกลับด้วยบัตรสมาชิก
// Reply API ฟรี ไม่เสียโควต้า push (ต่างจาก Push API ที่นับโควต้าต่อเดือน)
// ==============================================

/**
 * ส่ง Flex Message การ์ดสมาชิก (+ ข้อความ F) ผ่าน Reply API
 * @param {string} replyToken - replyToken จาก Webhook event (ใช้ได้ครั้งเดียว หมดอายุ ~1 นาที)
 * @param {Object} member     - { name, points, phone, expireDate }
 * @return {boolean} สำเร็จหรือไม่
 */
function replyMemberCardFlex(replyToken, member) {
  if (!replyToken) {
    Logger.log("⚠️ ไม่มี replyToken สำหรับส่ง Member Card");
    return false;
  }

  const messages = [{
    type: "flex",
    altText: "🎫 บัตรสมาชิกและคะแนนสะสมของคุณ",
    contents: buildMemberCardBubble(member)
  }];
  if (SEND_F_AFTER_REGISTER) {
    messages.push({ type: "text", text: F_MESSAGE_TEXT });
  }

  const url = "https://api.line.me/v2/bot/message/reply";
  const payload = {
    replyToken: replyToken,
    messages: messages
  };

  const options = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = response.getContentText();
    Logger.log("📬 Reply API Response: " + result);

    // Reply API สำเร็จจะตอบกลับ {} ส่วน error จะมี field "message"
    if (result.trim() !== "{}" && result.indexOf("\"message\"") !== -1) {
      Logger.log("❌ Reply API ตอบกลับ error: " + result);
      return false;
    }
    Logger.log("✅ ส่ง Member Card ผ่าน Reply API สำเร็จ (ฟรี ไม่เสียโควต้า)");
    return true;
  } catch (error) {
    Logger.log("❌ ส่ง Member Card ผ่าน Reply API ล้มเหลว: " + error.message);
    return false;
  }
}

/**
 * สร้าง Flex Bubble JSON สำหรับการ์ดสมาชิก
 * @param {Object} member - { name, points, phone, expireDate }
 */
function buildMemberCardBubble(member) {
  const name = member.name || "-";
  const points = member.points != null ? member.points : "-";
  const phone = member.phone || "-";
  const expireDate = member.expireDate || "-";

  const initial = String(name).trim().charAt(0).toUpperCase() || "M";

  return {
    type: "bubble",
    size: "mega",
    header: {
      type: "box",
      layout: "vertical",
      paddingAll: "20px",
      spacing: "md",
      background: {
        type: "linearGradient",
        angle: "135deg",
        startColor: "#1e1b4b",
        endColor: "#6d28d9"
      },
      contents: [
        {
          type: "box",
          layout: "horizontal",
          contents: [
            {
              type: "text",
              text: "🎫 MEMBER CARD",
              color: "#ffffff",
              weight: "bold",
              size: "sm",
              flex: 1
            },
            {
              type: "text",
              text: "★ PREMIUM",
              color: "#fbbf24",
              weight: "bold",
              size: "xxs",
              align: "end",
              gravity: "center"
            }
          ]
        },
        {
          type: "box",
          layout: "horizontal",
          spacing: "sm",
          margin: "md",
          contents: [
            {
              type: "box",
              layout: "vertical",
              width: "48px",
              height: "48px",
              cornerRadius: "24px",
              backgroundColor: "#fbbf24",
              contents: [
                {
                  type: "text",
                  text: initial,
                  size: "xl",
                  weight: "bold",
                  color: "#1e1b4b",
                  align: "center",
                  gravity: "center"
                }
              ]
            },
            {
              type: "text",
              text: name,
              color: "#ffffff",
              weight: "bold",
              size: "lg",
              gravity: "center",
              wrap: true,
              flex: 1
            }
          ]
        }
      ]
    },
    body: {
      type: "box",
      layout: "vertical",
      spacing: "md",
      paddingAll: "20px",
      contents: [
        {
          type: "box",
          layout: "horizontal",
          paddingAll: "14px",
          cornerRadius: "12px",
          background: {
            type: "linearGradient",
            angle: "90deg",
            startColor: "#fef3c7",
            endColor: "#fde68a"
          },
          contents: [
            {
              type: "box",
              layout: "vertical",
              flex: 1,
              contents: [
                {
                  type: "text",
                  text: "คะแนนสะสม",
                  size: "xs",
                  color: "#92400e"
                },
                {
                  type: "text",
                  text: String(points) + " คะแนน",
                  size: "xl",
                  weight: "bold",
                  color: "#78350f"
                }
              ]
            },
            {
              type: "text",
              text: "⭐",
              size: "xxl",
              gravity: "center"
            }
          ]
        },
        { type: "separator", margin: "md" },
        buildInfoRow("เบอร์โทร", phone),
        buildInfoRow("วันหมดอายุ", expireDate)
      ]
    },
    footer: {
      type: "box",
      layout: "vertical",
      paddingAll: "12px",
      contents: [
        {
          type: "text",
          text: "แสดงบัตรนี้ทุกครั้งเมื่อใช้บริการ 🛍️",
          size: "xs",
          color: "#9ca3af",
          align: "center",
          wrap: true
        }
      ]
    }
  };
}

/**
 * แถวข้อมูลแบบ label ซ้าย / value ขวา ใช้ในการ์ดสมาชิก
 */
function buildInfoRow(label, value) {
  return {
    type: "box",
    layout: "horizontal",
    contents: [
      { type: "text", text: label, size: "sm", color: "#888888", flex: 3 },
      { type: "text", text: String(value), size: "sm", color: "#1a1a1a", weight: "bold", flex: 4, align: "end", wrap: true }
    ]
  };
}

/**
 * แปลงวันที่เป็นรูปแบบไทย เช่น "31 ธ.ค. 2570"
 */
function formatThaiDate(date) {
  const thaiMonths = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
    "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
  const day = date.getDate();
  const month = thaiMonths[date.getMonth()];
  const buddhistYear = date.getFullYear() + 543;
  return day + " " + month + " " + buddhistYear;
}

// ==============================================
// ดึงข้อมูลโปรไฟล์ LINE OA ของระบบนี้ (ชื่อ + รูปโปรไฟล์)
// ==============================================

function getLineBotProfile() {
  const url = "https://api.line.me/v2/bot/info";
  const options = {
    method: "GET",
    headers: {
      "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const info = JSON.parse(response.getContentText());
    Logger.log("🏷️ ชื่อ OA: " + info.displayName);
    Logger.log("🖼️ pictureUrl: " + info.pictureUrl);
    Logger.log("🆔 basicId: " + info.basicId);
    return info;
  } catch (error) {
    Logger.log("❌ ดึงข้อมูล OA ล้มเหลว: " + error.message);
    return null;
  }
}

// ==============================================
// 🔗 LINE Login callback — แลก code → userId → Reply บัตรสมาชิก (ฟรี)
// เรียกจาก index.html หลังผู้ใช้กด "เชื่อมต่อกับ LINE" (LINE Login ทางการ)
// ==============================================

function handleOAuthCallback(params) {
  try {
    const code = params.code;
    const state = params.state;

    if (!LINE_LOGIN_CHANNEL_ID || LINE_LOGIN_CHANNEL_ID.indexOf("YOUR_") === 0) {
      return ContentService.createTextOutput("❌ ยังไม่ได้ตั้งค่า LINE_LOGIN_CHANNEL_ID / CHANNEL_SECRET");
    }

    // 1) แลก code → access token + id_token
    const tokenRes = UrlFetchApp.fetch("https://api.line.me/oauth2/v2.1/token", {
      method: "post",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      payload: {
        grant_type: "authorization_code",
        code: code,
        redirect_uri: FORM_REDIRECT_URI,
        client_id: LINE_LOGIN_CHANNEL_ID,
        client_secret: LINE_LOGIN_CHANNEL_SECRET
      },
      muteHttpExceptions: true
    });
    const tokenData = JSON.parse(tokenRes.getContentText());
    if (!tokenData.id_token) {
      Logger.log("❌ แลก token ไม่สำเร็จ: " + tokenRes.getContentText());
      return ContentService.createTextOutput("❌ เชื่อมต่อ LINE ไม่สำเร็จ: " + (tokenData.error_description || "เกิดข้อผิดพลาด"));
    }

    // 2) อ่าน userId จาก id_token (payload ส่วนกลาง)
    const idTokenPayload = JSON.parse(base64UrlDecode(tokenData.id_token.split(".")[1]));
    const userId = idTokenPayload.sub;
    Logger.log("🔗 LINE Login สำเร็จ userId: " + userId + " state: " + state);

    const props = PropertiesService.getScriptProperties();

    // 3) หาข้อมูลการลงทะเบียนที่รอเชื่อมต่อ (ผูกด้วย state)
    const pendingJson = props.getProperty("pending_" + state);
    if (!pendingJson) {
      return ContentService.createTextOutput("❌ ไม่พบข้อมูลการลงทะเบียนที่รอเชื่อมต่อ (state ไม่ตรง)");
    }
    const member = JSON.parse(pendingJson);
    props.deleteProperty("pending_" + state);

    // 4) ใช้ replyToken จาก link event (เก็บไว้ตอนเชื่อมต่อ) Reply บัตรสมาชิก (ฟรี ไม่เสียโควต้า)
    const savedToken = props.getProperty("token_" + userId);
    if (savedToken) {
      const sent = replyMemberCardFlex(savedToken, member);
      props.deleteProperty("token_" + userId);
      props.deleteProperty("ts_" + userId);
      if (sent) {
        props.deleteProperty("member_" + userId);
        return ContentService.createTextOutput("✅ เชื่อมต่อ LINE สำเร็จ — ส่งบัตรสมาชิกไปยัง LINE แล้ว (ฟรี ไม่เสียโควต้า)");
      }
      return ContentService.createTextOutput("⚠️ เชื่อมต่อ LINE สำเร็จ แต่ replyToken หมดอายุแล้ว — พิมพ์ข้อความในแชท LINE เพื่อรับบัตร");
    }

    // ไม่มี token (link event ยังมาไม่ถึง/หมดอายุ) → เก็บไว้รอ link event หรือพิมพ์ข้อความ (fallback)
    props.setProperty("pending_reply_" + userId, JSON.stringify(member));
    return ContentService.createTextOutput("✅ เชื่อมต่อ LINE สำเร็จ — กำลังส่งบัตรสมาชิก...");
  } catch (error) {
    Logger.log("❌ handleOAuthCallback Error: " + error.message);
    return ContentService.createTextOutput("❌ เกิดข้อผิดพลาด: " + error.message);
  }
}

/**
 * ถอดรหัส base64url (payload ของ id_token)
 */
function base64UrlDecode(str) {
  const b64 = str.replace(/-/g, "+").replace(/_/g, "/");
  const pad = b64.length % 4 === 0 ? "" : "=".repeat(4 - (b64.length % 4));
  return Utilities.newBlob(Utilities.base64Decode(b64 + pad)).getDataAsString();
}

// ==============================================
// getBotProfile - ดึงโปรไฟล์ LINE OA สำหรับหน้าเว็บ (fetch แบบ CORS)
// ใช้จาก index.html เพื่อดึง basicId สร้างลิงก์เปิดแชท LINE ทางการ
// โดยใช้ LINE_CHANNEL_ACCESS_TOKEN ที่ตั้งไว้ใน backend นี้โดยตรง
// ==============================================

function getBotProfile() {
  // ดึงจาก LINE API ทุกครั้ง (ไม่ cache) เพื่อให้หน้าเว็บ
  // แสดงชื่อ/โลโก้ตามบอทที่ใช้อยู่ล่าสุดเสมอ (AccessToken ใน backend นี้)
  const info = getLineBotProfile();
  if (!info) return { displayName: "", pictureUrl: "", basicId: "" };

  return {
    displayName: info.displayName || "",
    pictureUrl: info.pictureUrl || "",
    basicId: info.basicId || ""
  };
}

// ==============================================
// doGet - API
//   ?getProfile=1   → คืนค่า JSON โปรไฟล์ LINE OA (ใช้จากหน้า static ผ่าน CORS)
//   (ไม่มีพารามิเตอร์) → ดึงจำนวนข้อมูล (ใช้กับปุ่ม "ดูจำนวนผู้ลงทะเบียน")
// ==============================================

function doGet(e) {
  try {
    // ✅ LINE Login callback: แลก code → userId → Reply บัตรสมาชิก (ฟรี ไม่ใช้ push)
    if (e && e.parameter && e.parameter.completeLink === "1") {
      return handleOAuthCallback(e.parameter);
    }

    // API: ดึงโปรไฟล์ LINE OA (ตอบเป็น JSON สำหรับหน้า static)
    if (e && e.parameter && e.parameter.getProfile === "1") {
      const profile = getBotProfile();
      return ContentService.createTextOutput(JSON.stringify(profile))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // เดิม: ดึงจำนวนข้อมูล
    const lastRow = sheet.getLastRow();
    const count = Math.max(0, lastRow - 1);
    return ContentService.createTextOutput(String(count));
  } catch (error) {
    Logger.log("❌ doGet Error: " + error.message);
    return ContentService.createTextOutput("❌ " + error.message);
  }
}
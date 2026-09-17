import "dotenv/config";
import express from "express";
import nodemailer from "nodemailer";
import axios from "axios";
import mongoose from "mongoose";
import crypto from "crypto";
import UserDetails from "./_models/UserDetails.js";

// MongoDB connection (singleton, safe for serverless) -> database userDetailsGWC
let connectionPromise = null;

const connectDB = async () => {
  if (mongoose.connection.readyState === 1) return;
  if (!process.env.MONGODB_URI) throw new Error("MONGODB_URI is not set");
  if (!connectionPromise) {
    connectionPromise = mongoose
      .connect(process.env.MONGODB_URI, { dbName: "userDetailsGWC" })
      .then(() => console.log("Connected to MongoDB: userDetailsGWC"))
      .catch((err) => {
        connectionPromise = null;
        throw err;
      });
  }
  await connectionPromise;
};

// Configure Nodemailer transporter
const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST || "smtp.gmail.com",
  port: parseInt(process.env.SMTP_PORT || "587", 10),
  secure: process.env.SMTP_SECURE === "true",


  auth: {
    user: process.env.SMTP_USER || process.env.EMAIL_USER,
    pass: process.env.SMTP_PASS || process.env.EMAIL_PASS,
  },
});

// From = the user who submitted the form, To = fixed receiver
const TO_EMAIL = process.env.RECEIVER_EMAIL || "dhanushya.r@gwcdata.ai";
const CC_EMAILS = [
  "minithasri.k@gwcdata.ai",
  "aharsha.vhardhan@gwcdata.ai",
  "kaviya.priya@gwcdata.ai",
  "vishwanath.a@gwcdata.ai",
  "naraginti.chandu@gwcdata.ai",
];

async function sendEmailViaGraph(fromEmail, subject, emailBody, to, cc) {
  // 1. Get Access Token
  const tokenEndpoint = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
  const tokenParams = new URLSearchParams();
  tokenParams.append("client_id", process.env.CLIENT_ID);
  tokenParams.append("client_secret", process.env.CLIENT_SECRET);
  tokenParams.append("scope", "https://graph.microsoft.com/.default");
  tokenParams.append("grant_type", "client_credentials");

  const tokenResponse = await axios.post(tokenEndpoint, tokenParams, {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });

  const accessToken = tokenResponse.data.access_token;

  // 2. Send Email
  const graphEndpoint = `https://graph.microsoft.com/v1.0/users/${fromEmail}/sendMail`;

  const mailBody = {
    message: {
      subject,
      body: {
        contentType: "HTML",
        content: emailBody,
      },
      toRecipients: [
        {
          emailAddress: {
            address: to,
          },
        },
      ],
      ccRecipients: cc.map((address) => ({
        emailAddress: { address },
      })),
    },
    saveToSentItems: "true",
  };

  await axios.post(graphEndpoint, mailBody, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });
}

// Default recipients: To = TO_EMAIL, CC = team mails
async function sendEmail(
  fromEmail,
  subject,
  emailBody,
  { to = TO_EMAIL, cc = CC_EMAILS } = {},
) {
  const ccList = [...new Set(cc)].filter(
    (address) => address && address.toLowerCase() !== to.toLowerCase(),
  );

  if (process.env.CLIENT_ID && process.env.TENANT_ID) {
    console.log("Sending email via Microsoft Graph API...");
    await sendEmailViaGraph(fromEmail, subject, emailBody, to, ccList);
  } else {
    console.log("Sending email via Nodemailer...");
    await transporter.sendMail({
      from: fromEmail,
      to,
      cc: ccList.join(","),
      subject,
      html: emailBody,
    });
  }
}

const escapeHtml = (value) =>
  String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");

const formatDate = (value) => {
  if (!value) return "N/A";
  const date = new Date(value);
  return Number.isNaN(date.getTime())
    ? String(value)
    : date.toLocaleDateString("en-GB", {
        day: "2-digit",
        month: "short",
        year: "numeric",
        timeZone: "UTC",
      });
};

// Public base URL of this backend, used for the "Approve Request" link in emails.
// Set PUBLIC_API_URL when deployed; otherwise it is taken from the incoming request.
const getBaseUrl = (req) =>
  (
    process.env.PUBLIC_API_URL ||
    `${req.protocol}://${req.get("x-forwarded-host") || req.get("host")}`
  ).replace(/\/$/, "");

// approveUrl -> request email with an "Approve Request" button
// details.approvedBy -> approval email showing who approved it
function buildUserDetailsEmail(details, { approveUrl } = {}) {
  const row = (label, value, extraStyle = "") => `
            <tr>
              <td style="padding: 8px 0; font-weight: 600; color: #475569; width: 40%; vertical-align: top;">${label}</td>
              <td style="padding: 8px 0; color: #0f172a; ${extraStyle}">${escapeHtml(value || "N/A")}</td>
            </tr>`;

  const subdomain =
    details.subdomainRequired === "Yes"
      ? `Yes (${details.subdomainCount || "?"})`
      : details.subdomainRequired;

  const isApproved = details.approvalStatus === "Approved" && details.approvedBy;

  const heading = isApproved
    ? "Project Request Approved"
    : "New Project Request Submitted";

  const approvalSection = isApproved
    ? `
        <div style="background-color: #ecfdf5; border: 1px solid #a7f3d0; padding: 14px 20px; border-radius: 12px; margin-bottom: 24px; font-size: 14px; color: #065f46;">
          <strong>Approved By:</strong> ${escapeHtml(details.approvedBy)}
          ${details.approvedAt ? `<span style="color: #047857;"> &middot; ${formatDate(details.approvedAt)}</span>` : ""}
        </div>`
    : approveUrl
      ? `
        <div style="text-align: center; margin-bottom: 24px;">
          <p style="font-size: 13px; color: #64748b; margin: 0 0 12px 0;">Approver: click below and enter your name to approve this request.</p>
          <a href="${escapeHtml(approveUrl)}"
   style="background-color: #34A853; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 8px; font-weight: 600; font-size: 14px; display: inline-block;">
    Approve Request
</a>
        </div>`
      : "";

  return `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; padding: 24px; border: 1px solid #e2e8f0; border-radius: 16px; background-color: #ffffff; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);">
        <div style="border-bottom: 2px solid ${isApproved ? "#34A853" : "#4285F4"}; padding-bottom: 16px; margin-bottom: 24px;">
          <h2 style="color: ${isApproved ? "#188038" : "#1a73e8"}; margin: 0; font-size: 22px; font-weight: 700; letter-spacing: -0.02em;">${heading}</h2>
          <p style="margin: 4px 0 0 0; font-size: 13px; color: #5f6368;">GCP Project Request Intake System</p>
        </div>

        <div style="background-color: #f8fafc; border: 1px solid #f1f5f9; padding: 20px; border-radius: 12px; margin-bottom: 24px;">
          <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
            ${row("Requester Name", details.requesterName)}
            ${row("Project / POC Name", details.projectName)}
            ${row("Use Case", details.useCase)}
            ${row("Start Date", formatDate(details.startDate))}
            ${row("End Date", formatDate(details.endDate))}
            ${row("CI/CD Required", details.cicd)}
            ${row("Owner (Project / POC Manager)", details.ownerName)}
            ${row("RAM", details.ram)}
            ${row("Storage", details.storage)}
            ${row("Subdomain Needed", subdomain)}
            ${details.specification ? row("Additional Details", details.specification, "white-space: pre-wrap; line-height: 1.5;") : ""}
          </table>
        </div>
${approvalSection}
        <div style="text-align: center; margin-top: 24px; padding-top: 16px; border-top: 1px solid #f1f5f9;">
          <p style="font-size: 13px; color: #64748b; margin-bottom: 16px;">You can review this request and manage other submissions on the GCP dashboard page.</p>
          <a href="https://embed.domo.com/embed/pages/O8KrL"
   style="background-color: #0a0b0c; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 8px; font-weight: 600; font-size: 14px; display: inline-block; box-shadow: 0 4px 6px -1px rgba(66,133,244,0.25);">
    View Project Details Dashboard
</a>
        </div>
      </div>
    `;
}

// Small standalone HTML page for the approval link
function renderApprovalPage({ title, message = "", details, form = null, tone = "info" }) {
  const colors = {
    info: ["#1a73e8", "#e8f0fe"],
    success: ["#188038", "#e6f4ea"],
    error: ["#d93025", "#fce8e6"],
  }[tone];

  const summary = details
    ? `<table>
        <tr><td>Project / POC</td><td>${escapeHtml(details.projectName)}</td></tr>
        <tr><td>Requester</td><td>${escapeHtml(details.requesterName)}</td></tr>
        <tr><td>Owner</td><td>${escapeHtml(details.ownerName)}</td></tr>
        <tr><td>Duration</td><td>${formatDate(details.startDate)} &ndash; ${formatDate(details.endDate)}</td></tr>
      </table>`
    : "";

  const formHtml = form
    ? `<form method="POST" action="${escapeHtml(form.action)}">
        <input type="hidden" name="token" value="${escapeHtml(form.token)}" />
        <label for="approvedBy">Approved By</label>
        <input id="approvedBy" name="approvedBy" type="text" required maxlength="100"
               placeholder="Enter your full name" autofocus />
        <button type="submit">Approve Request</button>
      </form>`
    : "";

  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>${escapeHtml(title)}</title>
<style>
  body { margin: 0; min-height: 100vh; display: flex; align-items: center; justify-content: center;
         background: #f8fafc; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #0f172a; padding: 16px; box-sizing: border-box; }
  .card { width: 100%; max-width: 460px; background: #fff; border: 1px solid #e2e8f0; border-radius: 16px; padding: 28px; box-shadow: 0 10px 30px -5px rgba(30,41,59,.08); }
  h1 { margin: 0 0 8px; font-size: 22px; color: ${colors[0]}; }
  .msg { background: ${colors[1]}; color: ${colors[0]}; padding: 10px 14px; border-radius: 10px; font-size: 14px; margin: 12px 0; }
  table { width: 100%; border-collapse: collapse; font-size: 14px; margin: 16px 0; }
  td { padding: 6px 0; } td:first-child { color: #475569; font-weight: 600; width: 40%; }
  label { display: block; font-size: 13px; font-weight: 600; color: #475569; margin: 16px 0 6px; }
  input[type=text] { width: 100%; box-sizing: border-box; padding: 12px 14px; border: 1px solid #cbd5e1; border-radius: 10px; font-size: 14px; }
  input[type=text]:focus { outline: 2px solid #4285F4; border-color: transparent; }
  button { margin-top: 16px; width: 100%; padding: 12px; border: 0; border-radius: 10px; background: #34A853; color: #fff; font-size: 15px; font-weight: 600; cursor: pointer; }
  button:hover { background: #188038; }
</style>
</head>
<body>
  <div class="card">
    <h1>${escapeHtml(title)}</h1>
    ${message ? `<div class="msg">${escapeHtml(message)}</div>` : ""}
    ${summary}
    ${formHtml}
  </div>
</body>
</html>`;
}

const app = express();
app.set("trust proxy", true);
app.use(express.json());
app.use(express.urlencoded({ extended: false }));

app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader(
    "Access-Control-Allow-Methods",
    "GET, POST, OPTIONS, PUT, PATCH, DELETE",
  );
  res.setHeader(
    "Access-Control-Allow-Headers",
    "X-Requested-With,content-type,Authorization",
  );
  res.setHeader("Access-Control-Allow-Credentials", "true");
  if (req.method === "OPTIONS") {
    return res.sendStatus(200);
  }
  next();
});

// Save a user-details request to MongoDB (userDetailsGWC.requests), then email it
app.post("/api/requests", async (req, res) => {
  try {
    await connectDB();

    const {
      requesterName,
      projectName,
      useCase,
      startDate,
      endDate,
      cicd,
      ownerName,
      ram,
      storage,
      subdomainRequired,
      subdomainCount,
      specification,
      userEmail,
    } = req.body;

    if (startDate && endDate && new Date(endDate) < new Date(startDate)) {
      return res
        .status(400)
        .json({ error: "End date cannot be before start date" });
    }

    const approvalToken = crypto.randomBytes(24).toString("hex");

    const saved = await UserDetails.create({
      requesterName,
      projectName,
      useCase,
      startDate,
      endDate,
      cicd,
      ownerName,
      ram,
      storage,
      subdomainRequired,
      subdomainCount:
        subdomainRequired === "Yes" ? Number(subdomainCount) : undefined,
      specification,
      userEmail,
      approvalToken,
    });

    // Email failure should not lose the saved request
    // Sender: the Domo user, or SENDER_EMAIL when the form is opened outside Domo
    let emailSent = false;
    const fromEmail = userEmail || process.env.SENDER_EMAIL;
    if (fromEmail) {
      try {
        const approveUrl = `${getBaseUrl(req)}/api/requests/${saved._id}/approve?token=${approvalToken}`;
        await sendEmail(
          fromEmail,
          `New GCP Project Request: ${projectName || "Unknown"}`,
          buildUserDetailsEmail(saved, { approveUrl }),
        );
        emailSent = true;
      } catch (emailError) {
        console.error("Error sending email:", emailError);
      }
    }

    const request = saved.toObject();
    delete request.approvalToken;

    res.status(201).json({
      message: "Request saved successfully!",
      emailSent,
      request,
    });
  } catch (error) {
    console.error("Error saving request:", error);
    const status = error.name === "ValidationError" ? 400 : 500;
    res
      .status(status)
      .json({ error: "Failed to save request", details: error.message });
  }
});

// List saved user-details requests (newest first)
app.get("/api/requests", async (req, res) => {
  try {
    await connectDB();
    const requests = await UserDetails.find().sort({ createdAt: -1 });
    res.status(200).json(requests);
  } catch (error) {
    console.error("Error fetching requests:", error);
    res
      .status(500)
      .json({ error: "Failed to fetch requests", details: error.message });
  }
});

// Loads a request for the approval link, checking the id and token
const findForApproval = async (id, token) => {
  if (!mongoose.isValidObjectId(id) || !token) return null;
  const doc = await UserDetails.findById(id).select("+approvalToken");
  if (!doc || !doc.approvalToken) return null;
  const expected = Buffer.from(doc.approvalToken);
  const given = Buffer.from(String(token));
  if (expected.length !== given.length || !crypto.timingSafeEqual(expected, given)) {
    return null;
  }
  return doc;
};

const invalidLinkPage = () =>
  renderApprovalPage({
    title: "Invalid approval link",
    message: "This approval link is invalid or has expired.",
    tone: "error",
  });

const alreadyApprovedPage = (doc) =>
  renderApprovalPage({
    title: "Already approved",
    message: `This request was already approved by ${doc.approvedBy}.`,
    details: doc,
    tone: "success",
  });

// Approval page opened from the "Approve Request" button in the email
app.get("/api/requests/:id/approve", async (req, res) => {
  try {
    await connectDB();
    const doc = await findForApproval(req.params.id, req.query.token);
    if (!doc) return res.status(404).send(invalidLinkPage());
    if (doc.approvalStatus === "Approved") {
      return res.send(alreadyApprovedPage(doc));
    }

    res.send(
      renderApprovalPage({
        title: "Approve Project Request",
        details: doc,
        form: {
          action: `/api/requests/${doc._id}/approve`,
          token: req.query.token,
        },
      }),
    );
  } catch (error) {
    console.error("Error loading approval page:", error);
    res.status(500).send(
      renderApprovalPage({
        title: "Something went wrong",
        message: error.message,
        tone: "error",
      }),
    );
  }
});

// Save who approved the request, then send the approval email
app.post("/api/requests/:id/approve", async (req, res) => {
  try {
    await connectDB();
    const approvedBy = String(req.body.approvedBy || "").trim().slice(0, 100);
    const doc = await findForApproval(req.params.id, req.body.token);
    if (!doc) return res.status(404).send(invalidLinkPage());

    if (!approvedBy) {
      return res.status(400).send(
        renderApprovalPage({
          title: "Approve Project Request",
          message: "Please enter your name.",
          details: doc,
          tone: "error",
          form: {
            action: `/api/requests/${doc._id}/approve`,
            token: req.body.token,
          },
        }),
      );
    }

    // Only a pending request can be approved (prevents double approval)
    const approved = await UserDetails.findOneAndUpdate(
      { _id: doc._id, approvalStatus: { $ne: "Approved" } },
      {
        $set: {
          approvalStatus: "Approved",
          approvedBy,
          approvedAt: new Date(),
        },
      },
      { returnDocument: "after" },
    );
    if (!approved) {
      return res.send(alreadyApprovedPage(await UserDetails.findById(doc._id)));
    }

    // Approval email: To the requester, CC the receiver + team
    let emailNote = "An approval email has been sent to the requester and team.";
    const fromEmail =
      process.env.APPROVAL_FROM_EMAIL ||
      approved.userEmail ||
      process.env.SENDER_EMAIL;
    try {
      if (!fromEmail) throw new Error("No sender email available");
      await sendEmail(
        fromEmail,
        `Approved: GCP Project Request ${approved.projectName} (Approved by ${approvedBy})`,
        buildUserDetailsEmail(approved),
        {
          to: approved.userEmail || TO_EMAIL,
          cc: [TO_EMAIL, ...CC_EMAILS],
        },
      );
    } catch (emailError) {
      console.error("Error sending approval email:", emailError);
      emailNote = "Approval saved, but the approval email could not be sent.";
    }

    res.send(
      renderApprovalPage({
        title: "Request approved",
        message: `Approved by ${approvedBy}. ${emailNote}`,
        details: approved,
        tone: "success",
      }),
    );
  } catch (error) {
    console.error("Error approving request:", error);
    res.status(500).send(
      renderApprovalPage({
        title: "Something went wrong",
        message: error.message,
        tone: "error",
      }),
    );
  }
});

// Used by the project-details form: accepts a prebuilt { subject, body },
// or the user-details fields to build the body here.
app.post("/api/send-email", async (req, res) => {
  try {
    const { userEmail, subject: customSubject, body: customBody } = req.body;

    if (!userEmail) {
      return res.status(400).json({ error: "userEmail is required" });
    }

    const subject =
      customSubject ||
      `New GCP Project Request: ${req.body.projectName || "Unknown"}`;
    const emailBody = customBody || buildUserDetailsEmail(req.body);

    await sendEmail(userEmail, subject, emailBody);

    res.status(200).json({ message: "Email sent successfully!" });
  } catch (error) {
    console.error("Error sending email:", error);
    res
      .status(500)
      .json({ error: "Failed to send email", details: error.message });
  }
});

import { fileURLToPath } from "url";
import path from "path";
const currentFilePath = fileURLToPath(import.meta.url);
if (
  process.argv[1] &&
  path.resolve(process.argv[1]) === path.resolve(currentFilePath)
) {
  const PORT = process.env.PORT || 5000;
  app.listen(PORT, () => {
    console.log(`Backend server running on port ${PORT}`);
    connectDB().catch((err) => console.error("MongoDB connection error:", err));
  });
}

export default app;

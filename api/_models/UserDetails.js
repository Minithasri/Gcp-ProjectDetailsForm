import mongoose from "mongoose";

// Stored in MongoDB: userDetailsGWC -> requests
const userDetailsSchema = new mongoose.Schema(
  {
    requesterName: { type: String, required: true, trim: true },
    projectName: { type: String, required: true, trim: true },
    useCase: { type: String, required: true, trim: true },
    startDate: { type: Date, required: true },
    endDate: { type: Date, required: true },
    cicd: { type: String, enum: ["Yes", "No"], required: true },
    ownerName: { type: String, required: true, trim: true },
    // Added to the CC list of every mail for this request
    ownerEmail: {
      type: String,
      required: true,
      trim: true,
      lowercase: true,
      match: [/^[^\s@]+@[^\s@]+\.[^\s@]+$/, "Invalid owner email"],
    },
    ram: { type: String, required: true, trim: true },
    storage: { type: String, required: true, trim: true },
    subdomainRequired: { type: String, enum: ["Yes", "No"], required: true },
    subdomainCount: { type: Number, min: 1 },
    specification: { type: String, trim: true, default: "" },
    // Filled in later from the "Approve Request" link in the email
    approvalStatus: {
      type: String,
      enum: ["Pending", "Approved"],
      default: "Pending",
    },
    approvedBy: { type: String, trim: true, default: "" },
    approvedAt: { type: Date },
    approvalToken: { type: String, select: false },
    userEmail: { type: String, trim: true, default: "" },
    createdAt: { type: Date, default: Date.now },
  },
  { collection: "requests" },
);

export default mongoose.models.UserDetails ||
  mongoose.model("UserDetails", userDetailsSchema);

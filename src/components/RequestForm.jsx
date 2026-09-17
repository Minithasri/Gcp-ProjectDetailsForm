import React, { useState } from "react";
import domo from "ryuu.js";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "./ui/select";
import {
  Cloud,
  User,
  FolderKanban,
  Target,
  CalendarDays,
  CalendarCheck,
  GitBranch,
  UserCog,
  Cpu,
  HardDrive,
  Globe,
  Hash,
  FileText,
  Send,
  Loader2,
  CheckCircle2,
  AlertCircle,
} from "lucide-react";

// Email of the Domo user currently filling the form (empty outside Domo)
const getCurrentUserEmail = async () => {
  try {
    if (domo.env?.userEmail) return domo.env.userEmail;
    const env = await domo.get(`/domo/environment/v1/`);
    if (env?.userEmail) return env.userEmail;
    const user = await domo.get(
      `/domo/users/v1/${env?.userId || domo.env?.userId}?includeDetails=true`,
    );
    return user?.emailAddress || user?.email || user?.detail?.email || "";
  } catch (err) {
    console.warn("Could not resolve current user email:", err);
    return "";
  }
};

// Backend (project-request-app/api): saves to MongoDB userDetailsGWC.requests
// and sends the mail via Microsoft Graph. Empty = same origin (Vite dev proxy).
const API_URL = (
  import.meta.env.VITE_API_URL ||
  import.meta.env.VITE_EMAIL_API_URL ||
  ""
).replace(/\/$/, "");

const initialFormData = {
  requesterName: "",
  projectName: "",
  useCase: "",
  startDate: "",
  endDate: "",
  cicd: "",
  ownerName: "",
  ram: "",
  storage: "",
  subdomainRequired: "",
  subdomainCount: "",
  specification: "",
};

export const RequestForm = () => {
  const [formData, setFormData] = useState(initialFormData);

  const [status, setStatus] = useState({
    loading: false,
    message: "",
    type: "",
  });

  const handleChange = (e) => {
    setFormData({ ...formData, [e.target.name]: e.target.value });
  };

  const handleSelectChange = (name, value) => {
    setFormData((prev) => ({ ...prev, [name]: value }));
  };

  const handleSubmit = async (e) => {
    e.preventDefault();

    if (formData.endDate < formData.startDate) {
      setStatus({
        loading: false,
        message: "End date cannot be before start date.",
        type: "error",
      });
      return;
    }

    setStatus({ loading: true, message: "", type: "" });

    try {
      const userEmail = await getCurrentUserEmail();

      // Save to MongoDB (userDetailsGWC.requests); backend also sends the email
      const res = await fetch(`${API_URL}/api/requests`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...formData, userEmail }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) {
        throw new Error(data.details || data.error || `HTTP ${res.status}`);
      }
      console.log("User details saved to MongoDB:", data.request);

      setStatus({
        loading: false,
        message: data.emailSent
          ? "Request saved and email notification sent successfully!"
          : "Request saved successfully (email notification was not sent).",
        type: "success",
      });
      setFormData(initialFormData);
    } catch (error) {
      console.error(error);
      setStatus({
        loading: false,
        message: `Failed to submit request: ${error.message}`,
        type: "error",
      });
    }
  };

  const labelClass =
    "flex items-center gap-2 text-sm font-semibold text-gray-700 mb-2";
  const inputClass =
    "w-full px-4 py-3 rounded-xl border border-gray-200 bg-gray-50 text-gray-800 focus:bg-white focus:outline-none focus:ring-2 focus:ring-gcp-blue focus:border-transparent transition-all duration-200 appearance-none";
  const iconWrapClass =
    "w-5 h-5 rounded-md flex items-center justify-center shrink-0";
  const required = <span className="text-gcp-red">*</span>;

  const showSubdomainCount = formData.subdomainRequired === "Yes";

  return (
    <div className="min-h-screen bg-gray-50/30 flex items-center justify-center w-full p-4 sm:p-6 lg:p-10">
      <div className="w-full max-w-screen-2xl mx-auto bg-white shadow-xl shadow-gray-200/40 rounded-3xl border border-gray-100 overflow-hidden transition-all duration-300">
        {/* GCP four-color accent strip */}
        <div className="h-1.5 w-full flex">
          <div className="flex-1 bg-gcp-blue" />
          <div className="flex-1 bg-gcp-red" />
          <div className="flex-1 bg-gcp-yellow" />
          <div className="flex-1 bg-gcp-green" />
        </div>

        <div className="w-full px-6 py-10 sm:px-12 lg:px-16 xl:px-24">
          <div className="mb-10 border-b border-gray-100 pb-6 flex items-start gap-5">
            <div className="hidden sm:flex shrink-0 w-14 h-14 rounded-2xl bg-gcp-blue/10 items-center justify-center">
              <Cloud className="w-7 h-7 text-gcp-blue" strokeWidth={1.8} />
            </div>
            <div>
              <div className="flex items-center gap-3 mb-2">
                <h1 className="text-3xl sm:text-4xl lg:text-5xl font-extrabold text-gray-900 tracking-tight">
                  New Project Request
                </h1>
                <span className="hidden sm:inline-flex items-center px-3 py-1 rounded-full bg-gcp-blue/10 text-gcp-blue-dark text-xs font-bold tracking-wide">
                  GCP
                </span>
              </div>
              <p className="text-gray-500 text-lg sm:text-xl max-w-3xl">
                Submit your GCP project specifications below. Our engineering
                team will review and get back to you shortly.
              </p>
            </div>
          </div>

          {status.message && (
            <div
              className={`mb-8 p-4 rounded-xl flex items-center gap-3 font-medium animate-in fade-in slide-in-from-top-2 duration-300 ${
                status.type === "success"
                  ? "bg-gcp-green/10 text-gcp-green border border-gcp-green/30"
                  : "bg-gcp-red/10 text-gcp-red border border-gcp-red/30"
              }`}
            >
              {status.type === "success" ? (
                <CheckCircle2 className="w-5 h-5 shrink-0" />
              ) : (
                <AlertCircle className="w-5 h-5 shrink-0" />
              )}
              {status.message}
            </div>
          )}

          <form onSubmit={handleSubmit} className="space-y-8">
            {/* 1-2: Requester + Project / POC */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div>
                <label className={labelClass}>
                  <span className={`${iconWrapClass} bg-gcp-blue/10`}>
                    <User className="w-3.5 h-3.5 text-gcp-blue" />
                  </span>
                  Requester Name {required}
                </label>
                <input
                  type="text"
                  name="requesterName"
                  value={formData.requesterName}
                  onChange={handleChange}
                  placeholder="Your name"
                  required
                  className={inputClass}
                />
              </div>

              <div>
                <label className={labelClass}>
                  <span className={`${iconWrapClass} bg-gcp-red/10`}>
                    <FolderKanban className="w-3.5 h-3.5 text-gcp-red" />
                  </span>
                  Project / POC Name {required}
                </label>
                <input
                  type="text"
                  name="projectName"
                  value={formData.projectName}
                  onChange={handleChange}
                  placeholder="e.g. Nexus Platform"
                  required
                  className={inputClass}
                />
              </div>
            </div>

            {/* 3: Use case */}
            <div>
              <label className={labelClass}>
                <span className={`${iconWrapClass} bg-gcp-yellow/10`}>
                  <Target className="w-3.5 h-3.5 text-gcp-yellow" />
                </span>
                Use Case {required}
              </label>
              <input
                type="text"
                name="useCase"
                value={formData.useCase}
                onChange={handleChange}
                placeholder="e.g. Enterprise Tooling"
                required
                className={inputClass}
              />
            </div>

            {/* 4-7: Dates, CI/CD, Owner */}
            <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-6">
              <div>
                <label className={labelClass}>
                  <span className={`${iconWrapClass} bg-gcp-green/10`}>
                    <CalendarDays className="w-3.5 h-3.5 text-gcp-green" />
                  </span>
                  Start Date {required}
                </label>
                <input
                  type="date"
                  name="startDate"
                  value={formData.startDate}
                  onChange={handleChange}
                  required
                  className={inputClass}
                />
              </div>

              <div>
                <label className={labelClass}>
                  <span className={`${iconWrapClass} bg-gcp-blue/10`}>
                    <CalendarCheck className="w-3.5 h-3.5 text-gcp-blue" />
                  </span>
                  End Date {required}
                </label>
                <input
                  type="date"
                  name="endDate"
                  value={formData.endDate}
                  onChange={handleChange}
                  min={formData.startDate || undefined}
                  required
                  className={inputClass}
                />
              </div>

              <div>
                <label className={labelClass}>
                  <span className={`${iconWrapClass} bg-gcp-red/10`}>
                    <GitBranch className="w-3.5 h-3.5 text-gcp-red" />
                  </span>
                  CI/CD Required {required}
                </label>
                <Select
                  value={formData.cicd}
                  onValueChange={(val) => handleSelectChange("cicd", val)}
                  required
                >
                  <SelectTrigger className="w-full">
                    <SelectValue placeholder="Select option..." />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="Yes">Yes</SelectItem>
                    <SelectItem value="No">No</SelectItem>
                  </SelectContent>
                </Select>
              </div>

              <div>
                <label className={labelClass}>
                  <span className={`${iconWrapClass} bg-gcp-yellow/10`}>
                    <UserCog className="w-3.5 h-3.5 text-gcp-yellow" />
                  </span>
                  Owner Name {required}
                </label>
                <input
                  type="text"
                  name="ownerName"
                  value={formData.ownerName}
                  onChange={handleChange}
                  placeholder="Project / POC manager"
                  required
                  className={inputClass}
                />
              </div>
            </div>

            {/* 8: Detailed specification */}
            <div className="rounded-2xl border border-gray-100 bg-gray-50/50 p-5 sm:p-6 space-y-6">
              <div>
                <h2 className="flex items-center gap-2 text-base font-bold text-gray-800">
                  <FileText className="w-4 h-4 text-gcp-blue" />
                  Detailed Specification
                </h2>
                <p className="text-sm text-gray-500 mt-1">
                  Exact RAM and storage actually required by the project / POC,
                  and whether subdomains are needed.
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-6">
                <div>
                  <label className={labelClass}>
                    <span className={`${iconWrapClass} bg-gcp-blue/10`}>
                      <Cpu className="w-3.5 h-3.5 text-gcp-blue" />
                    </span>
                    RAM {required}
                  </label>
                  <input
                    type="text"
                    name="ram"
                    value={formData.ram}
                    onChange={handleChange}
                    placeholder="e.g. 8 GB"
                    required
                    className={inputClass}
                  />
                </div>

                <div>
                  <label className={labelClass}>
                    <span className={`${iconWrapClass} bg-gcp-green/10`}>
                      <HardDrive className="w-3.5 h-3.5 text-gcp-green" />
                    </span>
                    Storage {required}
                  </label>
                  <input
                    type="text"
                    name="storage"
                    value={formData.storage}
                    onChange={handleChange}
                    placeholder="e.g. 100 GB SSD"
                    required
                    className={inputClass}
                  />
                </div>

                <div>
                  <label className={labelClass}>
                    <span className={`${iconWrapClass} bg-gcp-red/10`}>
                      <Globe className="w-3.5 h-3.5 text-gcp-red" />
                    </span>
                    Subdomain Needed {required}
                  </label>
                  <Select
                    value={formData.subdomainRequired}
                    onValueChange={(val) =>
                      setFormData((prev) => ({
                        ...prev,
                        subdomainRequired: val,
                        subdomainCount: val === "Yes" ? prev.subdomainCount : "",
                      }))
                    }
                    required
                  >
                    <SelectTrigger className="w-full">
                      <SelectValue placeholder="Select option..." />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="Yes">Yes</SelectItem>
                      <SelectItem value="No">No</SelectItem>
                    </SelectContent>
                  </Select>
                </div>

                <div
                  className={`transition-opacity duration-300 ${showSubdomainCount ? "" : "opacity-50"}`}
                >
                  <label className={labelClass}>
                    <span className={`${iconWrapClass} bg-gcp-yellow/10`}>
                      <Hash className="w-3.5 h-3.5 text-gcp-yellow" />
                    </span>
                    No. of Subdomains {showSubdomainCount && required}
                  </label>
                  <input
                    type="number"
                    name="subdomainCount"
                    value={formData.subdomainCount}
                    onChange={handleChange}
                    placeholder={showSubdomainCount ? "e.g. 2" : "Select 'Yes' first"}
                    min="1"
                    disabled={!showSubdomainCount}
                    required={showSubdomainCount}
                    className={`${inputClass} ${showSubdomainCount ? "" : "cursor-not-allowed bg-gray-100"}`}
                  />
                </div>
              </div>

              <div>
                <label className={labelClass}>Additional Details</label>
                <textarea
                  name="specification"
                  value={formData.specification}
                  onChange={handleChange}
                  placeholder="CPU, OS, services, subdomain names, or anything else the team should know..."
                  className={`${inputClass} min-h-[120px] resize-y`}
                />
              </div>
            </div>

            {/* Submit action: bottom-right, compact icon button */}
            <div className="pt-4 border-t border-gray-100 flex justify-end">
              <button
                type="submit"
                disabled={status.loading}
                className="group inline-flex items-center gap-2.5 py-3.5 px-7 bg-gcp-blue hover:bg-gcp-blue-dark active:scale-[0.98] text-white text-base font-semibold rounded-xl shadow-lg shadow-gcp-blue/30 hover:shadow-gcp-blue/40 focus:ring-4 focus:ring-gcp-blue/50 transition-all duration-200 disabled:opacity-70 disabled:cursor-not-allowed disabled:active:scale-100"
              >
                {status.loading ? (
                  <>
                    Processing...
                    <Loader2 className="w-5 h-5 animate-spin" />
                  </>
                ) : (
                  <>
                    Submit Project Request
                    <Send className="w-4.5 h-4.5 transition-transform duration-200 group-hover:translate-x-0.5 group-hover:-translate-y-0.5" />
                  </>
                )}
              </button>
            </div>
          </form>
        </div>
      </div>
    </div>
  );
};

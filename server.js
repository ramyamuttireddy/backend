require("dotenv").config();

const express = require("express");
const cors = require("cors");
const { ClientSecretCredential } = require("@azure/identity");

const supabase = require("./config/supabase");

const app = express();

app.use(cors());
app.use(express.json());

// ======================================================
// AZURE
// ======================================================

const credential = new ClientSecretCredential(
  process.env.TENANT_ID,
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET
);

async function getAccessToken() {
  const tokenResponse = await credential.getToken(
    "https://graph.microsoft.com/.default"
  );

  return tokenResponse.token;
}

// ======================================================
// HOME
// ======================================================

app.get("/", (req, res) => {
  res.json({
    success: true,
    message: "Bharat Bhavan Salary API is running",
  });
});

// ======================================================
// EMPLOYEES FROM EXCEL
// ======================================================

app.get("/api/employees", async (req, res) => {
  try {
    console.log("Reading employees from Excel...");

    const userEmail = process.env.USER_EMAIL;
    const fileId = process.env.FILE_ID;
    const sheetName = process.env.EXCEL_SHEET_NAME;

    if (!userEmail) {
      return res.status(500).json({
        success: false,
        message: "USER_EMAIL missing in .env",
      });
    }

    if (!fileId) {
      return res.status(500).json({
        success: false,
        message: "FILE_ID missing in .env",
      });
    }

    if (!sheetName) {
      return res.status(500).json({
        success: false,
        message: "EXCEL_SHEET_NAME missing in .env",
      });
    }

    const accessToken = await getAccessToken();

    const graphUrl =
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
        userEmail
      )}/drive/items/${encodeURIComponent(
        fileId
      )}/workbook/worksheets/${encodeURIComponent(
        sheetName
      )}/usedRange`;

    console.log("Excel URL:");
    console.log(graphUrl);

    const response = await fetch(graphUrl, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
      },
    });

    const result = await response.json();

    if (!response.ok) {
      console.error("Graph Error:", result);

      return res.status(response.status).json({
        success: false,
        message:
          result?.error?.message ||
          "Failed to read Excel",
      });
    }

    const values = result.values || [];

    console.log("Excel total rows:", values.length);

    if (values.length === 0) {
      return res.json({
        success: true,
        source: "Excel",
        count: 0,
        data: [],
      });
    }

    const headers = values[0].map((header) =>
      String(header || "").trim()
    );

    console.log("Excel headers:");
    console.log(headers);

    const employees = values
      .slice(1)
      .map((row, index) => {
        const employee = {};

        headers.forEach((header, columnIndex) => {
          employee[header] =
            row[columnIndex] ?? "";
        });

        employee.id = index + 1;

        employee.name =
          employee["Full Name"] ||
          employee["Name"] ||
          employee["Employee Name"] ||
          "";

        employee.branch =
          employee["Location"] ||
          employee["Branch"] ||
          "";

        employee.email =
          employee["E-mail ID"] ||
          employee["Email"] ||
          employee["Email ID"] ||
          "";

        employee.phone =
          employee["Phone Number"] ||
          employee["Phone"] ||
          "";

        const salary =
          employee["Monthly Salary in Dollars"] ||
          employee["Monthly Salary"] ||
          0;

        employee.monthly_salary =
          Number(
            String(salary)
              .replace("$", "")
              .replace(",", "")
              .trim()
          ) || 0;

        return employee;
      })
      .filter((employee) =>
        String(employee.name).trim() !== ""
      );

    console.log(
      "Employees returned:",
      employees.length
    );

    return res.json({
      success: true,
      source: "Microsoft Excel",
      sheet: sheetName,
      count: employees.length,
      data: employees,
    });

  } catch (error) {
    console.error("Excel API Error:", error);

    return res.status(500).json({
      success: false,
      message: error.message,
    });
  }
});

// ======================================================
// ADD PAYMENT
// ======================================================

app.post("/salary/payments", async (req, res) => {
  try {
    console.log("================================");
    console.log("PAYMENT REQUEST");
    console.log(req.body);
    console.log("================================");

    const {
      employee_id,
      amount,
      payment_date,
      salary_month,
      remarks,
      notes,
    } = req.body;

    if (!employee_id) {
      return res.status(400).json({
        success: false,
        message: "Employee ID is required",
      });
    }

    if (!amount || Number(amount) <= 0) {
      return res.status(400).json({
        success: false,
        message: "Valid amount is required",
      });
    }

    if (!payment_date) {
      return res.status(400).json({
        success: false,
        message: "Payment date is required",
      });
    }

    // IMPORTANT
    const finalSalaryMonth =
      salary_month ||
      String(payment_date).substring(0, 7);

    const paymentData = {
      employee_id: Number(employee_id),
      amount: Number(amount),
      payment_date: payment_date,
      salary_month: finalSalaryMonth,
      remarks: remarks || null,
      notes: notes || null,
    };

    console.log("INSERT DATA:");
    console.log(paymentData);

    const { data, error } = await supabase
      .from("salary_payments")
      .insert(paymentData)
      .select()
      .single();

    if (error) {
      console.error("SUPABASE ERROR:");
      console.error(error);

      return res.status(500).json({
        success: false,
        message: error.message,
        details: error.details,
        code: error.code,
      });
    }

    console.log("PAYMENT SAVED:");
    console.log(data);

    return res.status(201).json({
      success: true,
      message: "Payment recorded successfully",
      payment: data,
    });

  } catch (error) {
    console.error("Payment Error:", error);

    return res.status(500).json({
      success: false,
      message: error.message,
    });
  }
});

// ======================================================
// GET PAYMENTS FOR EMPLOYEE
// ======================================================

app.get(
  "/salary/payments/:employeeId",
  async (req, res) => {
    try {
      const employeeId =
        Number(req.params.employeeId);

      if (!employeeId) {
        return res.status(400).json({
          success: false,
          message: "Employee ID is required",
        });
      }

      const { data, error } = await supabase
        .from("salary_payments")
        .select("*")
        .eq("employee_id", employeeId)
        .order("payment_date", {
          ascending: false,
        });

      if (error) {
        console.error(
          "Get payments error:",
          error
        );

        return res.status(500).json({
          success: false,
          message: error.message,
        });
      }

      return res.json({
        success: true,
        employee_id: employeeId,
        payments: data || [],
      });

    } catch (error) {
      return res.status(500).json({
        success: false,
        message: error.message,
      });
    }
  }
);


app.get("/salary/payments/:employeeId", async (req, res) => {
  try {
    const supabase = require("./config/supabase");

    const employeeId = Number(req.params.employeeId);
    const { month } = req.query;

    if (!employeeId) {
      return res.status(400).json({
        success: false,
        message: "Employee ID is required",
      });
    }

    if (!month) {
      return res.status(400).json({
        success: false,
        message: "Month is required. Example: 2026-08",
      });
    }

    const { data, error } = await supabase
      .from("salary_payments")
      .select("*")
      .eq("employee_id", employeeId)
      .eq("salary_month", month)
      .order("payment_date", {
        ascending: false,
      });

    if (error) {
      console.error("Get Payments Error:", error);

      return res.status(500).json({
        success: false,
        message: error.message,
      });
    }

    const totalPaid = (data || []).reduce(
      (total, payment) =>
        total + Number(payment.amount || 0),
      0
    );

    return res.json({
      success: true,
      employee_id: employeeId,
      salary_month: month,
      total_paid: totalPaid,
      payments: data || [],
    });

  } catch (error) {
    console.error("Payment API Error:", error);

    return res.status(500).json({
      success: false,
      message: error.message,
    });
  }
});

app.post("/salary/payments", async (req, res) => {
  try {
    const supabase = require("./config/supabase");

    const {
      employee_id,
      amount,
      payment_date,
      salary_month,
      notes,
    } = req.body;

    if (!employee_id) {
      return res.status(400).json({
        success: false,
        message: "Employee ID is required",
      });
    }

    if (!amount || Number(amount) <= 0) {
      return res.status(400).json({
        success: false,
        message: "Valid payment amount is required",
      });
    }

    if (!payment_date) {
      return res.status(400).json({
        success: false,
        message: "Payment date is required",
      });
    }

    // salary_month not provided -> derive from payment date
    const finalSalaryMonth =
      salary_month ||
      String(payment_date).slice(0, 7);

    const paymentData = {
      employee_id: Number(employee_id),
      amount: Number(amount),
      payment_date,
      salary_month: finalSalaryMonth,
      notes: notes || null,
    };

    console.log(
      "Saving payment:",
      paymentData
    );

    const { data, error } = await supabase
      .from("salary_payments")
      .insert([paymentData])
      .select()
      .single();

    if (error) {
      console.error(
        "Insert Payment Error:",
        error
      );

      return res.status(500).json({
        success: false,
        message: error.message,
      });
    }

    return res.status(201).json({
      success: true,
      message: "Payment recorded successfully",
      payment: data,
    });

  } catch (error) {
    console.error(
      "Payment API Error:",
      error
    );

    return res.status(500).json({
      success: false,
      message: error.message,
    });
  }
});

// ======================================================
// SERVER
// ======================================================

const PORT = process.env.PORT || 5000;

app.listen(PORT, () => {
  console.log(
    `Server running on port ${PORT}`
  );
});
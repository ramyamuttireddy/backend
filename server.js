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

// =====================================================
// UPDATE SALARY PAYMENT
// =====================================================

app.put("/salary/payments/:paymentId", async (req, res) => {
  try {
    const { paymentId } = req.params;

    const {
      amount,
      payment_date,
      salary_month,
      notes,
    } = req.body;

    if (!paymentId) {
      return res.status(400).json({
        success: false,
        message: "Payment ID is required",
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

    if (!salary_month) {
      return res.status(400).json({
        success: false,
        message: "Salary month is required",
      });
    }

    // -------------------------------------------------
    // Get existing payment
    // -------------------------------------------------

    const { data: existingPayment, error: existingError } =
      await supabase
        .from("salary_payments")
        .select("*")
        .eq("id", paymentId)
        .single();

    if (existingError) {
      console.error(existingError);

      return res.status(404).json({
        success: false,
        message: "Payment not found",
      });
    }

    // -------------------------------------------------
    // Get employee salary
    // -------------------------------------------------

    const employeeId = existingPayment.employee_id;

    const { data: salaryData, error: salaryError } =
      await supabase
        .from("salaries")
        .select("monthly_salary")
        .eq("employee_id", employeeId)
        .single();

    if (salaryError) {
      console.error(salaryError);

      return res.status(404).json({
        success: false,
        message: "Employee salary not found",
      });
    }

    const monthlySalary = Number(
      salaryData.monthly_salary || 0
    );

    // -------------------------------------------------
    // Calculate total paid for month
    //
    // IMPORTANT:
    // Exclude current payment because we are editing it
    // -------------------------------------------------

    const { data: monthPayments, error: paymentsError } =
      await supabase
        .from("salary_payments")
        .select("id, amount")
        .eq("employee_id", employeeId)
        .eq("salary_month", salary_month)
        .neq("id", paymentId);

    if (paymentsError) {
      console.error(paymentsError);

      return res.status(500).json({
        success: false,
        message: "Failed to calculate salary balance",
      });
    }

    const totalOtherPayments = (
      monthPayments || []
    ).reduce(
      (total, payment) =>
        total + Number(payment.amount || 0),
      0
    );

    const newAmount = Number(amount);

    const remainingBeforeThisPayment =
      monthlySalary - totalOtherPayments;

    if (newAmount > remainingBeforeThisPayment) {
      return res.status(400).json({
        success: false,
        message: `Payment cannot be greater than remaining salary ($${Math.max(
          remainingBeforeThisPayment,
          0
        ).toLocaleString()})`,
      });
    }

    // -------------------------------------------------
    // UPDATE PAYMENT
    // -------------------------------------------------

    const { data, error } = await supabase
      .from("salary_payments")
      .update({
        amount: newAmount,
        payment_date,
        salary_month,
        notes: notes || null,
      })
      .eq("id", paymentId)
      .select()
      .single();

    if (error) {
      console.error(error);

      return res.status(500).json({
        success: false,
        message: error.message,
      });
    }

    return res.json({
      success: true,
      message: "Payment updated successfully",
      payment: data,
    });
  } catch (error) {
    console.error("UPDATE PAYMENT ERROR:", error);

    return res.status(500).json({
      success: false,
      message: error.message || "Failed to update payment",
    });
  }
});

// =====================================================
// DELETE SALARY PAYMENT
// =====================================================

app.delete(
  "/salary/payments/:paymentId",
  async (req, res) => {
    try {
      const { paymentId } = req.params;

      if (!paymentId) {
        return res.status(400).json({
          success: false,
          message: "Payment ID is required",
        });
      }

      // -------------------------------------------------
      // Check payment exists
      // -------------------------------------------------

      const { data: payment, error: findError } =
        await supabase
          .from("salary_payments")
          .select("*")
          .eq("id", paymentId)
          .single();

      if (findError) {
        console.error(findError);

        return res.status(404).json({
          success: false,
          message: "Payment not found",
        });
      }

      // -------------------------------------------------
      // DELETE
      // -------------------------------------------------

      const { error } = await supabase
        .from("salary_payments")
        .delete()
        .eq("id", paymentId);

      if (error) {
        console.error(error);

        return res.status(500).json({
          success: false,
          message: error.message,
        });
      }

      return res.json({
        success: true,
        message: "Payment deleted successfully",
        deletedPayment: payment,
      });
    } catch (error) {
      console.error("DELETE PAYMENT ERROR:", error);

      return res.status(500).json({
        success: false,
        message:
          error.message || "Failed to delete payment",
      });
    }
  }
);



// ======================================================
// MONTHLY SALARY - GET
// ======================================================
// Example:
// GET /salary/monthly/1/2026-08
//
// This returns salary details + payments
// for ONE employee and ONE month.
// ======================================================

app.get(
  "/salary/monthly/:employeeId/:month",
  async (req, res) => {
    try {
      const employeeId = Number(
        req.params.employeeId
      );

      const month = req.params.month;

      if (!employeeId) {
        return res.status(400).json({
          success: false,
          message: "Employee ID is required",
        });
      }

      if (!month) {
        return res.status(400).json({
          success: false,
          message: "Salary month is required",
        });
      }

      // Convert:
      // 2026-08
      //
      // to:
      // 2026-08-01
      //
      // because salary_month column is DATE.

      const salaryMonth = `${month}-01`;

      // ------------------------------------------------
      // GET MONTHLY SALARY
      // ------------------------------------------------

      const {
        data: salaryData,
        error: salaryError,
      } = await supabase
        .from("monthly_salaries")
        .select("*")
        .eq("employee_id", employeeId)
        .eq("salary_month", salaryMonth)
        .maybeSingle();

      if (salaryError) {
        console.error(
          "Get monthly salary error:",
          salaryError
        );

        return res.status(500).json({
          success: false,
          message: salaryError.message,
        });
      }

      // ------------------------------------------------
      // GET PAYMENTS FOR SAME EMPLOYEE + SAME MONTH
      // ------------------------------------------------

      const {
        data: payments,
        error: paymentError,
      } = await supabase
        .from("salary_payments")
        .select("*")
        .eq("employee_id", employeeId)
        .eq("salary_month", month)
        .order("payment_date", {
          ascending: false,
        });

      if (paymentError) {
        console.error(
          "Get monthly payments error:",
          paymentError
        );

        return res.status(500).json({
          success: false,
          message: paymentError.message,
        });
      }

      // ------------------------------------------------
      // TOTAL PAID
      // ------------------------------------------------

      const totalPaid = (
        payments || []
      ).reduce(
        (total, payment) =>
          total + Number(payment.amount || 0),
        0
      );

      // ------------------------------------------------
      // IF NO SALARY EXISTS YET
      // ------------------------------------------------

      if (!salaryData) {
        return res.json({
          success: true,
          salary: null,
          payments: payments || [],
          total_paid: Number(
            totalPaid.toFixed(2)
          ),
          earned_salary: 0,
          remaining: 0,
        });
      }

      // ------------------------------------------------
      // CALCULATE EARNED SALARY
      // ------------------------------------------------

      const monthlySalary =
        Number(
          salaryData.monthly_salary || 0
        );

      const workingDays =
        Number(
          salaryData.working_days || 0
        );

      const hoursPerDay =
        Number(
          salaryData.hours_per_day || 0
        );

      const workedHours =
        Number(
          salaryData.worked_hours || 0
        );

      let earnedSalary = 0;

      if (
        monthlySalary > 0 &&
        workingDays > 0 &&
        hoursPerDay > 0
      ) {
        const totalMonthlyHours =
          workingDays *
          hoursPerDay;

        const hourlyRate =
          monthlySalary /
          totalMonthlyHours;

        earnedSalary =
          hourlyRate *
          workedHours;
      }

      earnedSalary = Number(
        earnedSalary.toFixed(2)
      );

      const remaining = Math.max(
        earnedSalary - totalPaid,
        0
      );

      return res.json({
        success: true,

        salary: {
          ...salaryData,

          earned_salary:
            earnedSalary,
        },

        payments:
          payments || [],

        total_paid:
          Number(
            totalPaid.toFixed(2)
          ),

        earned_salary:
          earnedSalary,

        remaining:
          Number(
            remaining.toFixed(2)
          ),
      });

    } catch (error) {
      console.error(
        "Monthly salary GET error:",
        error
      );

      return res.status(500).json({
        success: false,
        message: error.message,
      });
    }
  }
);


// ======================================================
// MONTHLY SALARY - SAVE
// ======================================================
// POST /salary/monthly
//
// Creates OR updates salary for:
// employee + month
// ======================================================

app.post(
  "/salary/monthly",
  async (req, res) => {
    try {
      const {
        employee_id,
        salary_month,
        monthly_salary,
        working_days,
        hours_per_day,
        worked_days,
        worked_hours,
      } = req.body;

      // ------------------------------------------------
      // VALIDATION
      // ------------------------------------------------

      if (!employee_id) {
        return res.status(400).json({
          success: false,
          message: "Employee ID is required",
        });
      }

      if (!salary_month) {
        return res.status(400).json({
          success: false,
          message: "Salary month is required",
        });
      }

      if (
        !monthly_salary ||
        Number(monthly_salary) <= 0
      ) {
        return res.status(400).json({
          success: false,
          message:
            "Valid monthly salary is required",
        });
      }

      if (
        !working_days ||
        Number(working_days) <= 0
      ) {
        return res.status(400).json({
          success: false,
          message:
            "Valid working days are required",
        });
      }

      if (
        !hours_per_day ||
        Number(hours_per_day) <= 0
      ) {
        return res.status(400).json({
          success: false,
          message:
            "Valid hours per day are required",
        });
      }

      // ------------------------------------------------
      // CONVERT MONTH
      // 2026-08 -> 2026-08-01
      // ------------------------------------------------

      const salaryMonthDate =
        `${salary_month}-01`;

      // ------------------------------------------------
      // CALCULATE EARNED SALARY
      // ------------------------------------------------

      const salary =
        Number(monthly_salary);

      const days =
        Number(working_days);

      const hours =
        Number(hours_per_day);

      const daysWorked =
        Number(worked_days || 0);

      let finalWorkedHours =
        Number(worked_hours || 0);

      // If worked hours is not entered,
      // calculate:
      //
      // worked days × hours per day

      if (
        finalWorkedHours <= 0 &&
        daysWorked > 0
      ) {
        finalWorkedHours =
          daysWorked * hours;
      }

      let earnedSalary = 0;

      if (
        salary > 0 &&
        days > 0 &&
        hours > 0
      ) {
        const totalMonthlyHours =
          days * hours;

        const hourlyRate =
          salary /
          totalMonthlyHours;

        earnedSalary =
          hourlyRate *
          finalWorkedHours;
      }

      earnedSalary = Number(
        earnedSalary.toFixed(2)
      );

      // ------------------------------------------------
      // UPSERT
      // ------------------------------------------------
      // employee_id + salary_month is UNIQUE
      // so same month updates existing record.
      // ------------------------------------------------

      const salaryData = {
        employee_id:
          Number(employee_id),

        salary_month:
          salaryMonthDate,

        monthly_salary:
          salary,

        working_days:
          days,

        hours_per_day:
          hours,

        worked_days:
          daysWorked,

        worked_hours:
          finalWorkedHours,

        earned_salary:
          earnedSalary,

        updated_at:
          new Date().toISOString(),
      };

      const {
        data,
        error,
      } = await supabase
        .from("monthly_salaries")
        .upsert(
          salaryData,
          {
            onConflict:
              "employee_id,salary_month",
          }
        )
        .select()
        .single();

      if (error) {
        console.error(
          "Save monthly salary error:",
          error
        );

        return res.status(500).json({
          success: false,
          message: error.message,
          details: error.details,
          code: error.code,
        });
      }

      return res.status(201).json({
        success: true,
        message:
          "Monthly salary saved successfully",

        salary: data,
      });

    } catch (error) {
      console.error(
        "Monthly salary POST error:",
        error
      );

      return res.status(500).json({
        success: false,
        message: error.message,
      });
    }
  }
);


// ======================================================
// MONTHLY SALARY - UPDATE
// ======================================================
// PUT /salary/monthly/:employeeId/:month
// ======================================================

app.put(
  "/salary/monthly/:employeeId/:month",
  async (req, res) => {
    try {
      const employeeId =
        Number(
          req.params.employeeId
        );

      const month =
        req.params.month;

      const {
        monthly_salary,
        working_days,
        hours_per_day,
        worked_days,
        worked_hours,
      } = req.body;

      if (!employeeId) {
        return res.status(400).json({
          success: false,
          message:
            "Employee ID is required",
        });
      }

      if (!month) {
        return res.status(400).json({
          success: false,
          message:
            "Salary month is required",
        });
      }

      // ------------------------------------------------
      // CALCULATE
      // ------------------------------------------------

      const salary =
        Number(monthly_salary || 0);

      const days =
        Number(working_days || 0);

      const hours =
        Number(hours_per_day || 0);

      const daysWorked =
        Number(worked_days || 0);

      let finalWorkedHours =
        Number(worked_hours || 0);

      if (
        finalWorkedHours <= 0 &&
        daysWorked > 0
      ) {
        finalWorkedHours =
          daysWorked * hours;
      }

      let earnedSalary = 0;

      if (
        salary > 0 &&
        days > 0 &&
        hours > 0
      ) {
        const totalMonthlyHours =
          days * hours;

        const hourlyRate =
          salary /
          totalMonthlyHours;

        earnedSalary =
          hourlyRate *
          finalWorkedHours;
      }

      earnedSalary = Number(
        earnedSalary.toFixed(2)
      );

      const salaryMonthDate =
        `${month}-01`;

      // ------------------------------------------------
      // UPDATE
      // ------------------------------------------------

      const {
        data,
        error,
      } = await supabase
        .from("monthly_salaries")
        .update({
          monthly_salary:
            salary,

          working_days:
            days,

          hours_per_day:
            hours,

          worked_days:
            daysWorked,

          worked_hours:
            finalWorkedHours,

          earned_salary:
            earnedSalary,

          updated_at:
            new Date().toISOString(),
        })
        .eq(
          "employee_id",
          employeeId
        )
        .eq(
          "salary_month",
          salaryMonthDate
        )
        .select()
        .single();

      if (error) {
        console.error(
          "Update monthly salary error:",
          error
        );

        return res.status(500).json({
          success: false,
          message: error.message,
        });
      }

      return res.json({
        success: true,

        message:
          "Monthly salary updated successfully",

        salary: data,
      });

    } catch (error) {
      console.error(
        "Monthly salary PUT error:",
        error
      );

      return res.status(500).json({
        success: false,
        message: error.message,
      });
    }
  }
);




// ======================================================
// SERVER
// ======================================================

const PORT = process.env.PORT || 5000;

app.listen(PORT, () => {
  console.log(
    `Server running on port ${PORT}`
  );
});
require("dotenv").config();

const express = require("express");
const cors = require("cors");

const {
  ClientSecretCredential,
} = require("@azure/identity");

const supabase = require("./config/supabase");

const app = express();

/* =========================================================
   MIDDLEWARE
========================================================= */

app.use(
  cors({
    origin: "*",
    methods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);

app.use(express.json());

/* =========================================================
   ENVIRONMENT CHECK
========================================================= */

const requiredEnv = [
  "SUPABASE_URL",
  "SUPABASE_SERVICE_ROLE_KEY",
  "TENANT_ID",
  "CLIENT_ID",
  "CLIENT_SECRET",
  "USER_EMAIL",
  "FILE_ID",
  "EXCEL_SHEET_NAME",
];

for (const key of requiredEnv) {
  if (!process.env[key]) {
    console.warn(`WARNING: ${key} is missing`);
  }
}

/* =========================================================
   AZURE
========================================================= */

const credential = new ClientSecretCredential(
  process.env.TENANT_ID,
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET
);

/* =========================================================
   TOKEN CACHE
========================================================= */

let cachedAccessToken = null;
let tokenExpiresAt = 0;

async function getAccessToken() {
  const now = Date.now();

  if (
    cachedAccessToken &&
    now < tokenExpiresAt
  ) {
    return cachedAccessToken;
  }

  console.log("Getting new Microsoft Graph token...");

  const tokenResponse =
    await credential.getToken(
      "https://graph.microsoft.com/.default"
    );

  if (!tokenResponse || !tokenResponse.token) {
    throw new Error(
      "Unable to get Microsoft Graph access token"
    );
  }

  cachedAccessToken =
    tokenResponse.token;

  tokenExpiresAt =
    tokenResponse.expiresOnTimestamp -
    5 * 60 * 1000;

  return cachedAccessToken;
}

/* =========================================================
   EMPLOYEE CACHE
========================================================= */

let employeesCache = null;
let employeesCacheTime = 0;

const EMPLOYEE_CACHE_TIME =
  5 * 60 * 1000;

/* =========================================================
   HELPERS
========================================================= */

function normalizeMonth(month) {
  if (!month) {
    return null;
  }

  const value =
    String(month).trim();

  if (/^\d{4}-\d{2}$/.test(value)) {
    return value;
  }

  if (
    /^\d{4}-\d{2}-\d{2}$/.test(value)
  ) {
    return value.slice(0, 7);
  }

  return null;
}

function monthToDate(month) {
  const normalizedMonth =
    normalizeMonth(month);

  if (!normalizedMonth) {
    return null;
  }

  return `${normalizedMonth}-01`;
}

function toNumber(value) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return 0;
  }

  const number =
    Number(
      String(value)
        .replace("$", "")
        .replace(",", "")
        .trim()
    );

  return Number.isFinite(number)
    ? number
    : 0;
}

function round2(value) {
  return Number(
    Number(value || 0).toFixed(2)
  );
}

/* =========================================================
   FETCH EMPLOYEES FROM EXCEL
========================================================= */

async function fetchEmployeesFromExcel() {
  const now = Date.now();

  if (
    employeesCache &&
    now - employeesCacheTime <
      EMPLOYEE_CACHE_TIME
  ) {
    console.log(
      "Returning employees from cache"
    );

    return employeesCache;
  }

  console.log(
    "Reading employees from Excel..."
  );

  const userEmail =
    process.env.USER_EMAIL;

  const fileId =
    process.env.FILE_ID;

  const sheetName =
    process.env.EXCEL_SHEET_NAME;

  if (!userEmail) {
    throw new Error(
      "USER_EMAIL missing in environment variables"
    );
  }

  if (!fileId) {
    throw new Error(
      "FILE_ID missing in environment variables"
    );
  }

  if (!sheetName) {
    throw new Error(
      "EXCEL_SHEET_NAME missing in environment variables"
    );
  }

  const accessToken =
    await getAccessToken();

  const graphUrl =
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      userEmail
    )}/drive/items/${encodeURIComponent(
      fileId
    )}/workbook/worksheets/${encodeURIComponent(
      sheetName
    )}/usedRange`;

  const response =
    await fetch(graphUrl, {
      method: "GET",
      headers: {
        Authorization:
          `Bearer ${accessToken}`,
        Accept:
          "application/json",
      },
    });

  const result =
    await response.json();

  if (!response.ok) {
    console.error(
      "Microsoft Graph Error:",
      result
    );

    throw new Error(
      result?.error?.message ||
        "Failed to read Excel"
    );
  }

  const values =
    result.values || [];

  if (!Array.isArray(values) ||
      values.length === 0) {

    employeesCache = [];
    employeesCacheTime =
      Date.now();

    return [];
  }

  const headers =
    values[0].map((header) =>
      String(
        header || ""
      ).trim()
    );

  const employees =
    values
      .slice(1)
      .map(
        (row, index) => {
          const employee = {};

          headers.forEach(
            (
              header,
              columnIndex
            ) => {
              employee[header] =
                row[columnIndex] ??
                "";
            }
          );

          /*
             Keep the Excel row based ID
          */

          employee.id =
            index + 1;

          /*
             Name
          */

          employee.name =
            employee["Full Name"] ||
            employee["Name"] ||
            employee["Employee Name"] ||
            "";

          /*
             Branch
          */

          employee.branch =
            employee["Location"] ||
            employee["Branch"] ||
            "";

          /*
             Email
          */

          employee.email =
            employee["E-mail ID"] ||
            employee["Email"] ||
            employee["Email ID"] ||
            "";

          /*
             Phone
          */

          employee.phone =
            employee["Phone Number"] ||
            employee["Phone"] ||
            "";

          /*
             Monthly Salary
          */

          const salary =
            employee[
              "Monthly Salary in Dollars"
            ] ||
            employee[
              "Monthly Salary"
            ] ||
            0;

          employee.monthly_salary =
            toNumber(salary);

          return employee;
        }
      )
      .filter(
        (employee) =>
          String(
            employee.name
          ).trim() !== ""
      );

  employeesCache =
    employees;

  employeesCacheTime =
    Date.now();

  return employees;
}

/* =========================================================
   HOME
========================================================= */

app.get(
  "/",
  (req, res) => {
    res.json({
      success: true,
      message:
        "Bharat Bhavan Salary API is running",
    });
  }
);

/* =========================================================
   HEALTH CHECK
========================================================= */

app.get(
  "/health",
  async (req, res) => {
    try {
      const {
        error,
      } = await supabase
        .from("monthly_salaries")
        .select("employee_id")
        .limit(1);

      if (error) {
        return res.status(500).json({
          success: false,
          supabase: false,
          message:
            error.message,
        });
      }

      return res.json({
        success: true,
        supabase: true,
        message:
          "API and Supabase are working",
      });
    } catch (error) {
      return res.status(500).json({
        success: false,
        supabase: false,
        message:
          error.message,
      });
    }
  }
);

/* =========================================================
   GET EMPLOYEES
========================================================= */

app.get(
  "/api/employees",
  async (req, res) => {
    try {
      const employees =
        await fetchEmployeesFromExcel();

      return res.json({
        success: true,
        source:
          "Microsoft Excel",
        count:
          employees.length,
        data:
          employees,
      });
    } catch (error) {
      console.error(
        "Employees API Error:",
        error
      );

      return res.status(500).json({
        success: false,
        message:
          error.message,
      });
    }
  }
);

/* =========================================================
   EMPLOYEE SUMMARY
========================================================= */

app.get(
  "/api/employees/summary",
  async (req, res) => {
    try {
      const employees =
        await fetchEmployeesFromExcel();

      let friscoCount = 0;
      let irvingCount = 0;

      let friscoSalary = 0;
      let irvingSalary = 0;

      employees.forEach(
        (employee) => {
          const branch =
            String(
              employee.branch || ""
            )
              .trim()
              .toLowerCase();

          const salary =
            toNumber(
              employee.monthly_salary
            );

          if (
            branch.includes("frisco")
          ) {
            friscoCount++;
            friscoSalary += salary;
          }

          if (
            branch.includes("irving")
          ) {
            irvingCount++;
            irvingSalary += salary;
          }
        }
      );

      return res.json({
        success: true,

        frisco: {
          employees:
            friscoCount,
          salary:
            round2(friscoSalary),
        },

        irving: {
          employees:
            irvingCount,
          salary:
            round2(irvingSalary),
        },

        total: {
          employees:
            employees.length,
          salary:
            round2(
              friscoSalary +
              irvingSalary
            ),
        },
      });
    } catch (error) {
      console.error(
        "Employee summary error:",
        error
      );

      return res.status(500).json({
        success: false,
        message:
          error.message,
      });
    }
  }
);

/* =========================================================
   GET ALL MONTHLY SALARIES
========================================================= */

app.get(
  "/salary/monthly",
  async (req, res) => {
    try {
      const month =
        normalizeMonth(
          req.query.month
        );

      if (!month) {
        return res.status(400).json({
          success: false,
          message:
            "Valid month is required. Example: 2026-08",
        });
      }

      const salaryMonthDate =
        monthToDate(month);

      const {
        data,
        error,
      } =
        await supabase
          .from(
            "monthly_salaries"
          )
          .select("*")
          .eq(
            "salary_month",
            salaryMonthDate
          )
          .order(
            "employee_id"
          );

      if (error) {
        return res.status(500).json({
          success: false,
          message:
            error.message,
        });
      }

      return res.json({
        success: true,
        salary_month:
          month,
        salaries:
          data || [],
      });
    } catch (error) {
      console.error(
        "Monthly salaries error:",
        error
      );

      return res.status(500).json({
        success: false,
        message:
          error.message,
      });
    }
  }
);

/* =========================================================
   GET MONTHLY SALARY + PAYMENTS
========================================================= */

app.get(
  "/salary/monthly/:employeeId/:month",
  async (req, res) => {
    try {
      const employeeId =
        Number(
          req.params.employeeId
        );

      const month =
        normalizeMonth(
          req.params.month
        );

      if (
        !employeeId ||
        !Number.isFinite(employeeId)
      ) {
        return res.status(400).json({
          success: false,
          message:
            "Valid employee ID is required",
        });
      }

      if (!month) {
        return res.status(400).json({
          success: false,
          message:
            "Valid salary month is required",
        });
      }

      const salaryMonthDate =
        monthToDate(month);

      /*
         Salary and payments are fetched
         simultaneously.
      */

      const [
        salaryResult,
        paymentResult,
      ] = await Promise.all([
        supabase
          .from(
            "monthly_salaries"
          )
          .select("*")
          .eq(
            "employee_id",
            employeeId
          )
          .eq(
            "salary_month",
            salaryMonthDate
          )
          .maybeSingle(),

        supabase
          .from(
            "salary_payments"
          )
          .select("*")
          .eq(
            "employee_id",
            employeeId
          )
          .eq(
            "salary_month",
            month
          )
          .order(
            "payment_date",
            {
              ascending: false,
            }
          ),
      ]);

      const {
        data: salaryData,
        error: salaryError,
      } = salaryResult;

      const {
        data: payments,
        error: paymentError,
      } = paymentResult;

      if (salaryError) {
        return res.status(500).json({
          success: false,
          message:
            salaryError.message,
        });
      }

      if (paymentError) {
        return res.status(500).json({
          success: false,
          message:
            paymentError.message,
        });
      }

      const totalPaid =
        (payments || []).reduce(
          (
            total,
            payment
          ) =>
            total +
            toNumber(
              payment.amount
            ),
          0
        );

      /*
         No monthly salary saved yet
      */

      if (!salaryData) {
        return res.json({
          success: true,
          salary: null,
          payments:
            payments || [],
          total_paid:
            round2(totalPaid),
          earned_salary: 0,
          remaining: 0,
        });
      }

      const monthlySalary =
        toNumber(
          salaryData.monthly_salary
        );

      const workingDays =
        toNumber(
          salaryData.working_days
        );

      const hoursPerDay =
        toNumber(
          salaryData.hours_per_day
        );

      const workedHours =
        toNumber(
          salaryData.worked_hours
        );

      let earnedSalary =
        toNumber(
          salaryData.earned_salary
        );

      /*
         Recalculate earned salary
      */

      if (
        monthlySalary > 0 &&
        workingDays > 0 &&
        hoursPerDay > 0
      ) {
        earnedSalary =
          (
            monthlySalary /
            (
              workingDays *
              hoursPerDay
            )
          ) *
          workedHours;
      }

      earnedSalary =
        round2(
          earnedSalary
        );

      const remaining =
        Math.max(
          earnedSalary -
            totalPaid,
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
          round2(totalPaid),

        earned_salary:
          earnedSalary,

        remaining:
          round2(remaining),
      });
    } catch (error) {
      console.error(
        "Monthly salary error:",
        error
      );

      return res.status(500).json({
        success: false,
        message:
          error.message,
      });
    }
  }
);

/* =========================================================
   SAVE MONTHLY SALARY
========================================================= */

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

      if (!employee_id) {
        return res.status(400).json({
          success: false,
          message:
            "Employee ID is required",
        });
      }

      const month =
        normalizeMonth(
          salary_month
        );

      if (!month) {
        return res.status(400).json({
          success: false,
          message:
            "Valid salary month is required. Example: 2026-08",
        });
      }

      const salary =
        toNumber(
          monthly_salary
        );

      const days =
        toNumber(
          working_days
        );

      const hours =
        toNumber(
          hours_per_day
        );

      const daysWorked =
        toNumber(
          worked_days
        );

      let workedHours =
        toNumber(
          worked_hours
        );

      if (
        workedHours <= 0 &&
        daysWorked > 0
      ) {
        workedHours =
          daysWorked *
          hours;
      }

      let earnedSalary = 0;

      if (
        salary > 0 &&
        days > 0 &&
        hours > 0
      ) {
        earnedSalary =
          (
            salary /
            (days * hours)
          ) *
          workedHours;
      }

      earnedSalary =
        round2(
          earnedSalary
        );

      const salaryData = {
        employee_id:
          Number(employee_id),

        salary_month:
          monthToDate(month),

        monthly_salary:
          salary,

        working_days:
          days,

        hours_per_day:
          hours,

        worked_days:
          daysWorked,

        worked_hours:
          workedHours,

        earned_salary:
          earnedSalary,

        updated_at:
          new Date().toISOString(),
      };

      const {
        data,
        error,
      } =
        await supabase
          .from(
            "monthly_salaries"
          )
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
          "Save salary error:",
          error
        );

        return res.status(500).json({
          success: false,
          message:
            error.message,
        });
      }

      return res.status(201).json({
        success: true,
        message:
          "Monthly salary saved successfully",
        salary:
          data,
      });
    } catch (error) {
      console.error(
        "Save monthly salary error:",
        error
      );

      return res.status(500).json({
        success: false,
        message:
          error.message,
      });
    }
  }
);

/* =========================================================
   UPDATE MONTHLY SALARY
========================================================= */

app.put(
  "/salary/monthly/:employeeId/:month",
  async (req, res) => {
    try {
      const employeeId =
        Number(
          req.params.employeeId
        );

      const month =
        normalizeMonth(
          req.params.month
        );

      if (
        !employeeId ||
        !Number.isFinite(employeeId)
      ) {
        return res.status(400).json({
          success: false,
          message:
            "Valid employee ID is required",
        });
      }

      if (!month) {
        return res.status(400).json({
          success: false,
          message:
            "Valid salary month is required",
        });
      }

      const {
        monthly_salary,
        working_days,
        hours_per_day,
        worked_days,
        worked_hours,
      } = req.body;

      const salary =
        toNumber(
          monthly_salary
        );

      const days =
        toNumber(
          working_days
        );

      const hours =
        toNumber(
          hours_per_day
        );

      const daysWorked =
        toNumber(
          worked_days
        );

      let workedHours =
        toNumber(
          worked_hours
        );

      if (
        workedHours <= 0 &&
        daysWorked > 0
      ) {
        workedHours =
          daysWorked *
          hours;
      }

      let earnedSalary = 0;

      if (
        salary > 0 &&
        days > 0 &&
        hours > 0
      ) {
        earnedSalary =
          (
            salary /
            (days * hours)
          ) *
          workedHours;
      }

      earnedSalary =
        round2(
          earnedSalary
        );

      const {
        data,
        error,
      } =
        await supabase
          .from(
            "monthly_salaries"
          )
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
              workedHours,

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
            monthToDate(month)
          )
          .select()
          .single();

      if (error) {
        return res.status(500).json({
          success: false,
          message:
            error.message,
        });
      }

      return res.json({
        success: true,
        message:
          "Monthly salary updated successfully",
        salary:
          data,
      });
    } catch (error) {
      console.error(
        "Update monthly salary error:",
        error
      );

      return res.status(500).json({
        success: false,
        message:
          error.message,
      });
    }
  }
);

/* =========================================================
   ADD PAYMENT
========================================================= */

app.post(
  "/salary/payments",
  async (req, res) => {
    try {
      const {
        employee_id,
        amount,
        payment_date,
        salary_month,
        notes,
      } = req.body;

      const employeeId =
        Number(employee_id);

      const newAmount =
        toNumber(amount);

      if (
        !employeeId ||
        !Number.isFinite(employeeId)
      ) {
        return res.status(400).json({
          success: false,
          message:
            "Valid employee ID is required",
        });
      }

      if (newAmount <= 0) {
        return res.status(400).json({
          success: false,
          message:
            "Valid payment amount is required",
        });
      }

      if (!payment_date) {
        return res.status(400).json({
          success: false,
          message:
            "Payment date is required",
        });
      }

      const finalSalaryMonth =
        normalizeMonth(
          salary_month
        ) ||
        String(
          payment_date
        ).slice(0, 7);

      if (!finalSalaryMonth) {
        return res.status(400).json({
          success: false,
          message:
            "Valid salary month is required",
        });
      }

      /*
         Get salary and existing payments
         together.
      */

      const [
        salaryResult,
        paymentsResult,
      ] = await Promise.all([
        supabase
          .from(
            "monthly_salaries"
          )
          .select(
            "earned_salary, monthly_salary"
          )
          .eq(
            "employee_id",
            employeeId
          )
          .eq(
            "salary_month",
            monthToDate(
              finalSalaryMonth
            )
          )
          .maybeSingle(),

        supabase
          .from(
            "salary_payments"
          )
          .select("amount")
          .eq(
            "employee_id",
            employeeId
          )
          .eq(
            "salary_month",
            finalSalaryMonth
          ),
      ]);

      if (
        salaryResult.error
      ) {
        return res.status(500).json({
          success: false,
          message:
            salaryResult.error.message,
        });
      }

      if (
        paymentsResult.error
      ) {
        return res.status(500).json({
          success: false,
          message:
            paymentsResult.error.message,
        });
      }

      /*
         Prefer earned salary.
         If not available, use monthly salary.
      */

      const salaryLimit =
        toNumber(
          salaryResult.data
            ?.earned_salary
        ) ||
        toNumber(
          salaryResult.data
            ?.monthly_salary
        );

      const alreadyPaid =
        (
          paymentsResult.data ||
          []
        ).reduce(
          (
            total,
            payment
          ) =>
            total +
            toNumber(
              payment.amount
            ),
          0
        );

      const remaining =
        Math.max(
          salaryLimit -
            alreadyPaid,
          0
        );

      /*
         Do not allow payment greater
         than remaining salary.
      */

      if (
        salaryLimit > 0 &&
        newAmount > remaining
      ) {
        return res.status(400).json({
          success: false,
          message:
            `Payment exceeds remaining salary. Remaining: $${remaining.toFixed(
              2
            )}`,
        });
      }

      const {
        data,
        error,
      } =
        await supabase
          .from(
            "salary_payments"
          )
          .insert({
            employee_id:
              employeeId,

            amount:
              newAmount,

            payment_date,

            salary_month:
              finalSalaryMonth,

            notes:
              notes || null,
          })
          .select()
          .single();

      if (error) {
        return res.status(500).json({
          success: false,
          message:
            error.message,
        });
      }

      return res.status(201).json({
        success: true,
        message:
          "Payment recorded successfully",
        payment:
          data,
      });
    } catch (error) {
      console.error(
        "Add payment error:",
        error
      );

      return res.status(500).json({
        success: false,
        message:
          error.message,
      });
    }
  }
);

/* =========================================================
   GET PAYMENTS FOR EMPLOYEE
========================================================= */

app.get(
  "/salary/payments/:employeeId",
  async (req, res) => {
    try {
      const employeeId =
        Number(
          req.params.employeeId
        );

      if (
        !employeeId ||
        !Number.isFinite(employeeId)
      ) {
        return res.status(400).json({
          success: false,
          message:
            "Valid employee ID is required",
        });
      }

      const {
        data,
        error,
      } =
        await supabase
          .from(
            "salary_payments"
          )
          .select("*")
          .eq(
            "employee_id",
            employeeId
          )
          .order(
            "payment_date",
            {
              ascending: false,
            }
          );

      if (error) {
        return res.status(500).json({
          success: false,
          message:
            error.message,
        });
      }

      return res.json({
        success: true,
        employee_id:
          employeeId,
        payments:
          data || [],
      });
    } catch (error) {
      console.error(
        "Get payments error:",
        error
      );

      return res.status(500).json({
        success: false,
        message:
          error.message,
      });
    }
  }
);

/* =========================================================
   UPDATE PAYMENT
========================================================= */

app.put(
  "/salary/payments/:paymentId",
  async (req, res) => {
    try {
      const paymentId =
        req.params.paymentId;

      const {
        amount,
        payment_date,
        salary_month,
        notes,
      } = req.body;

      const newAmount =
        toNumber(amount);

      if (newAmount <= 0) {
        return res.status(400).json({
          success: false,
          message:
            "Valid payment amount is required",
        });
      }

      if (!payment_date) {
        return res.status(400).json({
          success: false,
          message:
            "Payment date is required",
        });
      }

      const {
        data: existingPayment,
        error: existingError,
      } =
        await supabase
          .from(
            "salary_payments"
          )
          .select("*")
          .eq(
            "id",
            paymentId
          )
          .single();

      if (
        existingError ||
        !existingPayment
      ) {
        return res.status(404).json({
          success: false,
          message:
            "Payment not found",
        });
      }

      const employeeId =
        Number(
          existingPayment.employee_id
        );

      /*
         If salary_month isn't supplied,
         use the existing payment month.
      */

      const finalSalaryMonth =
        normalizeMonth(
          salary_month
        ) ||
        normalizeMonth(
          existingPayment.salary_month
        );

      if (!finalSalaryMonth) {
        return res.status(400).json({
          success: false,
          message:
            "Valid salary month is required",
        });
      }

      const {
        data: salaryData,
        error: salaryError,
      } =
        await supabase
          .from(
            "monthly_salaries"
          )
          .select(
            "monthly_salary, earned_salary"
          )
          .eq(
            "employee_id",
            employeeId
          )
          .eq(
            "salary_month",
            monthToDate(
              finalSalaryMonth
            )
          )
          .maybeSingle();

      if (salaryError) {
        return res.status(500).json({
          success: false,
          message:
            salaryError.message,
        });
      }

      const salaryLimit =
        toNumber(
          salaryData?.earned_salary
        ) ||
        toNumber(
          salaryData?.monthly_salary
        );

      /*
         Get all OTHER payments.
      */

      const {
        data: otherPayments,
        error: paymentsError,
      } =
        await supabase
          .from(
            "salary_payments"
          )
          .select(
            "id, amount"
          )
          .eq(
            "employee_id",
            employeeId
          )
          .eq(
            "salary_month",
            finalSalaryMonth
          )
          .neq(
            "id",
            paymentId
          );

      if (paymentsError) {
        return res.status(500).json({
          success: false,
          message:
            paymentsError.message,
        });
      }

      const totalOther =
        (
          otherPayments ||
          []
        ).reduce(
          (
            total,
            payment
          ) =>
            total +
            toNumber(
              payment.amount
            ),
          0
        );

      const remainingForPayment =
        Math.max(
          salaryLimit -
            totalOther,
          0
        );

      if (
        salaryLimit > 0 &&
        newAmount >
          remainingForPayment
      ) {
        return res.status(400).json({
          success: false,
          message:
            `Payment is greater than remaining salary. Maximum allowed: $${remainingForPayment.toFixed(
              2
            )}`,
        });
      }

      const {
        data,
        error,
      } =
        await supabase
          .from(
            "salary_payments"
          )
          .update({
            amount:
              newAmount,

            payment_date,

            salary_month:
              finalSalaryMonth,

            notes:
              notes || null,
          })
          .eq(
            "id",
            paymentId
          )
          .select()
          .single();

      if (error) {
        return res.status(500).json({
          success: false,
          message:
            error.message,
        });
      }

      return res.json({
        success: true,
        message:
          "Payment updated successfully",
        payment:
          data,
      });
    } catch (error) {
      console.error(
        "Update payment error:",
        error
      );

      return res.status(500).json({
        success: false,
        message:
          error.message,
      });
    }
  }
);

/* =========================================================
   DELETE PAYMENT
========================================================= */

app.delete(
  "/salary/payments/:paymentId",
  async (req, res) => {
    try {
      const paymentId =
        req.params.paymentId;

      const {
        data: payment,
        error: findError,
      } =
        await supabase
          .from(
            "salary_payments"
          )
          .select("*")
          .eq(
            "id",
            paymentId
          )
          .single();

      if (
        findError ||
        !payment
      ) {
        return res.status(404).json({
          success: false,
          message:
            "Payment not found",
        });
      }

      const {
        error,
      } =
        await supabase
          .from(
            "salary_payments"
          )
          .delete()
          .eq(
            "id",
            paymentId
          );

      if (error) {
        return res.status(500).json({
          success: false,
          message:
            error.message,
        });
      }

      return res.json({
        success: true,
        message:
          "Payment deleted successfully",
        deletedPayment:
          payment,
      });
    } catch (error) {
      console.error(
        "Delete payment error:",
        error
      );

      return res.status(500).json({
        success: false,
        message:
          error.message,
      });
    }
  }
);

/* =========================================================
   404 HANDLER
========================================================= */

app.use(
  (req, res) => {
    res.status(404).json({
      success: false,
      message:
        "API endpoint not found",
      path:
        req.originalUrl,
    });
  }
);

/* =========================================================
   GLOBAL ERROR HANDLER
========================================================= */

app.use(
  (
    error,
    req,
    res,
    next
  ) => {
    console.error(
      "Unhandled server error:",
      error
    );

    res.status(500).json({
      success: false,
      message:
        error.message ||
        "Internal server error",
    });
  }
);

/* =========================================================
   SERVER
========================================================= */

const PORT =
  process.env.PORT || 5000;

app.listen(
  PORT,
  () => {
    console.log(
      "======================================"
    );

    console.log(
      `Bharat Bhavan Salary API running on port ${PORT}`
    );

    console.log(
      `http://localhost:${PORT}`
    );

    console.log(
      "======================================"
    );
  }
);
require("dotenv").config();

const express = require("express");
const cors = require("cors");

const {
  ClientSecretCredential,
} = require("@azure/identity");

const supabase = require("./config/supabase");

const app = express();

// ======================================================
// MIDDLEWARE
// ======================================================

app.use(
  cors({
    origin: "*",
  })
);

app.use(express.json());

// ======================================================
// AZURE
// ======================================================

const credential =
  new ClientSecretCredential(
    process.env.TENANT_ID,
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET
  );

// ======================================================
// TOKEN CACHE
// ======================================================

let cachedAccessToken = null;
let tokenExpiresAt = 0;

async function getAccessToken() {
  const now = Date.now();

  // Reuse token if still valid
  if (
    cachedAccessToken &&
    now < tokenExpiresAt
  ) {
    return cachedAccessToken;
  }

  console.log(
    "Getting new Microsoft Graph token..."
  );

  const tokenResponse =
    await credential.getToken(
      "https://graph.microsoft.com/.default"
    );

  cachedAccessToken =
    tokenResponse.token;

  // Keep 5 minutes safety buffer
  tokenExpiresAt =
    tokenResponse.expiresOnTimestamp -
    5 * 60 * 1000;

  return cachedAccessToken;
}

// ======================================================
// EMPLOYEE CACHE
// ======================================================

let employeesCache = null;
let employeesCacheTime = 0;

const EMPLOYEE_CACHE_TIME =
  5 * 60 * 1000;

// ======================================================
// GET EMPLOYEES FROM EXCEL
// ======================================================

async function fetchEmployeesFromExcel() {
  const now = Date.now();

  // Return cached data
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
      "USER_EMAIL missing in .env"
    );
  }

  if (!fileId) {
    throw new Error(
      "FILE_ID missing in .env"
    );
  }

  if (!sheetName) {
    throw new Error(
      "EXCEL_SHEET_NAME missing in .env"
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
      "Graph Error:",
      result
    );

    throw new Error(
      result?.error?.message ||
        "Failed to read Excel"
    );
  }

  const values =
    result.values || [];

  if (values.length === 0) {
    employeesCache = [];

    employeesCacheTime = Date.now();

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

          employee.id =
            index + 1;

          employee.name =
            employee[
              "Full Name"
            ] ||
            employee[
              "Name"
            ] ||
            employee[
              "Employee Name"
            ] ||
            "";

          employee.branch =
            employee[
              "Location"
            ] ||
            employee[
              "Branch"
            ] ||
            "";

          employee.email =
            employee[
              "E-mail ID"
            ] ||
            employee[
              "Email"
            ] ||
            employee[
              "Email ID"
            ] ||
            "";

          employee.phone =
            employee[
              "Phone Number"
            ] ||
            employee[
              "Phone"
            ] ||
            "";

          const salary =
            employee[
              "Monthly Salary in Dollars"
            ] ||
            employee[
              "Monthly Salary"
            ] ||
            0;

          employee.monthly_salary =
            Number(
              String(salary)
                .replace("$", "")
                .replace(",", "")
                .trim()
            ) || 0;

          return employee;
        }
      )
      .filter(
        (employee) =>
          String(
            employee.name
          ).trim() !== ""
      );

  // Save cache
  employeesCache =
    employees;

  employeesCacheTime =
    Date.now();

  return employees;
}

// ======================================================
// HOME
// ======================================================

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

// ======================================================
// GET EMPLOYEES
// ======================================================

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

// ======================================================
// EMPLOYEE SUMMARY
// ======================================================

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
            Number(
              employee.monthly_salary ||
                0
            );

          if (
            branch.includes(
              "frisco"
            )
          ) {
            friscoCount++;

            friscoSalary +=
              salary;
          }

          if (
            branch.includes(
              "irving"
            )
          ) {
            irvingCount++;

            irvingSalary +=
              salary;
          }
        }
      );

      return res.json({
        success: true,

        frisco: {
          employees:
            friscoCount,

          salary:
            Number(
              friscoSalary.toFixed(
                2
              )
            ),
        },

        irving: {
          employees:
            irvingCount,

          salary:
            Number(
              irvingSalary.toFixed(
                2
              )
            ),
        },

        total: {
          employees:
            employees.length,

          salary:
            Number(
              (
                friscoSalary +
                irvingSalary
              ).toFixed(2)
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

// ======================================================
// GET ALL MONTHLY SALARIES
// ======================================================

app.get(
  "/salary/monthly",
  async (req, res) => {
    try {
      const month =
        req.query.month;

      if (!month) {
        return res.status(400).json({
          success: false,

          message:
            "Month is required. Example: 2026-08",
        });
      }

      const salaryMonthDate =
        `${month}-01`;

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
      return res.status(500).json({
        success: false,

        message:
          error.message,
      });
    }
  }
);

// ======================================================
// GET MONTHLY SALARY + PAYMENTS
// ======================================================

app.get(
  "/salary/monthly/:employeeId/:month",
  async (req, res) => {
    try {
      const employeeId =
        Number(
          req.params.employeeId
        );

      const month =
        req.params.month;

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

      const salaryMonthDate =
        `${month}-01`;

      // Salary + payments in parallel
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
              ascending:
                false,
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
            Number(
              payment.amount ||
                0
            ),
          0
        );

      if (!salaryData) {
        return res.json({
          success: true,

          salary: null,

          payments:
            payments || [],

          total_paid:
            Number(
              totalPaid.toFixed(
                2
              )
            ),

          earned_salary: 0,

          remaining: 0,
        });
      }

      const monthlySalary =
        Number(
          salaryData.monthly_salary ||
            0
        );

      const workingDays =
        Number(
          salaryData.working_days ||
            0
        );

      const hoursPerDay =
        Number(
          salaryData.hours_per_day ||
            0
        );

      const workedHours =
        Number(
          salaryData.worked_hours ||
            0
        );

      let earnedSalary =
        Number(
          salaryData.earned_salary ||
            0
        );

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
        Number(
          earnedSalary.toFixed(
            2
          )
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
          Number(
            totalPaid.toFixed(
              2
            )
          ),

        earned_salary:
          earnedSalary,

        remaining:
          Number(
            remaining.toFixed(
              2
            )
          ),
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

// ======================================================
// SAVE MONTHLY SALARY
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

      if (!employee_id) {
        return res.status(400).json({
          success: false,

          message:
            "Employee ID is required",
        });
      }

      if (!salary_month) {
        return res.status(400).json({
          success: false,

          message:
            "Salary month is required",
        });
      }

      const salary =
        Number(
          monthly_salary || 0
        );

      const days =
        Number(
          working_days || 0
        );

      const hours =
        Number(
          hours_per_day || 0
        );

      const daysWorked =
        Number(
          worked_days || 0
        );

      let workedHours =
        Number(
          worked_hours || 0
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
        Number(
          earnedSalary.toFixed(
            2
          )
        );

      const salaryData = {
        employee_id:
          Number(employee_id),

        salary_month:
          `${salary_month}-01`,

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
      return res.status(500).json({
        success: false,

        message:
          error.message,
      });
    }
  }
);

// ======================================================
// UPDATE MONTHLY SALARY
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

      const salary =
        Number(
          monthly_salary || 0
        );

      const days =
        Number(
          working_days || 0
        );

      const hours =
        Number(
          hours_per_day || 0
        );

      const daysWorked =
        Number(
          worked_days || 0
        );

      let workedHours =
        Number(
          worked_hours || 0
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
        Number(
          earnedSalary.toFixed(
            2
          )
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
            `${month}-01`
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
      return res.status(500).json({
        success: false,

        message:
          error.message,
      });
    }
  }
);

// ======================================================
// ADD PAYMENT
// ======================================================

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

      if (!employee_id) {
        return res.status(400).json({
          success: false,

          message:
            "Employee ID is required",
        });
      }

      if (
        !amount ||
        Number(amount) <= 0
      ) {
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
        salary_month ||
        String(
          payment_date
        ).slice(0, 7);

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
            Number(employee_id)
          )
          .eq(
            "salary_month",
            `${finalSalaryMonth}-01`
          )
          .maybeSingle(),

        supabase
          .from(
            "salary_payments"
          )
          .select("amount")
          .eq(
            "employee_id",
            Number(employee_id)
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
            salaryResult.error
              .message,
        });
      }

      if (
        paymentsResult.error
      ) {
        return res.status(500).json({
          success: false,

          message:
            paymentsResult.error
              .message,
        });
      }

      const salaryLimit =
        Number(
          salaryResult.data
            ?.earned_salary ||
            salaryResult.data
              ?.monthly_salary ||
            0
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
            Number(
              payment.amount ||
                0
            ),
          0
        );

      const newAmount =
        Number(amount);

      if (
        salaryLimit > 0 &&
        alreadyPaid +
          newAmount >
          salaryLimit
      ) {
        return res.status(400).json({
          success: false,

          message:
            `Payment exceeds remaining salary. Remaining: $${Math.max(
              salaryLimit -
                alreadyPaid,
              0
            ).toFixed(2)}`,
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
              Number(employee_id),

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
      return res.status(500).json({
        success: false,

        message:
          error.message,
      });
    }
  }
);

// ======================================================
// GET PAYMENTS
// ======================================================

app.get(
  "/salary/payments/:employeeId",
  async (req, res) => {
    try {
      const employeeId =
        Number(
          req.params.employeeId
        );

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
              ascending:
                false,
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
      return res.status(500).json({
        success: false,

        message:
          error.message,
      });
    }
  }
);

// ======================================================
// UPDATE PAYMENT
// ======================================================

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

      if (existingError) {
        return res.status(404).json({
          success: false,

          message:
            "Payment not found",
        });
      }

      const employeeId =
        existingPayment.employee_id;

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
            `${salary_month}-01`
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
        Number(
          salaryData?.earned_salary ||
            salaryData?.monthly_salary ||
            0
        );

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
            salary_month
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
            Number(
              payment.amount ||
                0
            ),
          0
        );

      const newAmount =
        Number(amount);

      if (
        salaryLimit > 0 &&
        newAmount >
          Math.max(
            salaryLimit -
              totalOther,
            0
          )
      ) {
        return res.status(400).json({
          success: false,

          message:
            "Payment is greater than remaining salary",
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

            salary_month,

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
      return res.status(500).json({
        success: false,

        message:
          error.message,
      });
    }
  }
);

// ======================================================
// DELETE PAYMENT
// ======================================================

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

      if (findError) {
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
      return res.status(500).json({
        success: false,

        message:
          error.message,
      });
    }
  }
);

// ======================================================
// SERVER
// ======================================================

const PORT =
  process.env.PORT || 5000;

app.listen(
  PORT,
  () => {
    console.log(
      `Server running on port ${PORT}`
    );
  }
);
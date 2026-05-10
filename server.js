require("dotenv").config();

const express = require("express");
const mongoose = require("mongoose");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");
const cors = require("cors");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const ExcelJS = require("exceljs");
const cron = require("node-cron");

const app = express();

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use(express.static("public"));
app.get("/", (req, res) => {
  res.sendFile(__dirname + "/public/login.html");
});
app.use("/uploads", express.static("uploads"));

/* ---------------- MONGODB ---------------- */

mongoose
  .connect(process.env.MONGO_URI)
  .then(() => {
    console.log("MongoDB Connected");
  })
  .catch((err) => {
    console.log(err);
  });

/* ---------------- MODELS ---------------- */

const userSchema = new mongoose.Schema({
  email: String,
  password: String,
});

const customerSchema = new mongoose.Schema({
  firstName: String,
  lastName: String,
  email: String,
  phone: String,
  city: String,
  address: String,
  notes: String,
  files: [String],
  createdAt: {
    type: Date,
    default: Date.now,
  },
});

const User = mongoose.model("User", userSchema);
const Customer = mongoose.model("Customer", customerSchema);

/* ---------------- MULTER ---------------- */

if (!fs.existsSync("uploads")) {
  fs.mkdirSync("uploads");
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, "uploads/");
  },

  filename: (req, file, cb) => {
    cb(
      null,
      Date.now() + "-" + file.originalname
    );
  },
});

const upload = multer({ storage });

/* ---------------- AUTH ---------------- */

app.post("/register", async (req, res) => {
  try {
    const { email, password } = req.body;

    const existingUser =
      await User.findOne({ email });

    if (existingUser) {
      return res.json({
        success: false,
        message: "User already exists",
      });
    }

    const hashedPassword =
      await bcrypt.hash(password, 10);

    const user = new User({
      email,
      password: hashedPassword,
    });

    await user.save();

    res.json({
      success: true,
    });
  } catch (err) {
    console.log(err);

    res.json({
      success: false,
      message: "Register Failed",
    });
  }
});

app.post("/login", async (req, res) => {
  try {
    const { email, password } = req.body;

    const user =
      await User.findOne({ email });

    if (!user) {
      return res.json({
        success: false,
        message: "Invalid Credentials",
      });
    }

    const validPassword =
      await bcrypt.compare(
        password,
        user.password
      );

    if (!validPassword) {
      return res.json({
        success: false,
        message: "Invalid Credentials",
      });
    }

    const token = jwt.sign(
      {
        id: user._id,
      },
      process.env.JWT_SECRET
    );

    res.json({
      success: true,
      token,
    });
  } catch (err) {
    console.log(err);

    res.json({
      success: false,
      message: "Login Failed",
    });
  }
});

/* ---------------- CUSTOMERS ---------------- */

app.post(
  "/add-customer",
  upload.array("files"),
  async (req, res) => {
    try {
      const {
        firstName,
        lastName,
        email,
        phone,
        city,
        address,
        notes,
      } = req.body;

      if (
        !firstName ||
        !lastName ||
        !email ||
        !city
      ) {
        return res.json({
          success: false,
          message:
            "Please fill required fields",
        });
      }

      const cleanPhone =
  phone.replace(/\D/g,'');

if(cleanPhone.length < 7){

  return res.json({

    success:false,

    message:
      "Phone number is invalid"

  });

}

      const files = req.files.map(
        (file) => file.filename
      );

      const customer = new Customer({
        firstName,
        lastName,
        email,
        phone,
        city,
        address,
        notes,
        files,
      });

      await customer.save();

      res.json({
        success: true,
      });
    } catch (err) {
      console.log(err);

      res.json({
        success: false,
        message: "Failed",
      });
    }
  }
);

app.get("/customers", async (req, res) => {
  try {
    const customers =
      await Customer.find().sort({
        createdAt: -1,
      });

    res.json(customers);
  } catch (err) {
    console.log(err);
  }
});

app.delete(
  "/delete-customer/:id",
  async (req, res) => {
    try {
      await Customer.findByIdAndDelete(
        req.params.id
      );

      res.json({
        success: true,
      });
    } catch (err) {
      console.log(err);
    }
  }
);

/* ---------------- BACKUP ---------------- */

async function createBackup() {
  try {
    const customers =
      await Customer.find();

    const workbook =
      new ExcelJS.Workbook();

    const worksheet =
      workbook.addWorksheet("Customers");

    worksheet.columns = [
      {
        header: "First Name",
        key: "firstName",
        width: 20,
      },
      {
        header: "Last Name",
        key: "lastName",
        width: 20,
      },
      {
        header: "Email",
        key: "email",
        width: 30,
      },
      {
        header: "Phone",
        key: "phone",
        width: 20,
      },
      {
        header: "City",
        key: "city",
        width: 20,
      },
      {
        header: "Address",
        key: "address",
        width: 40,
      },
      {
        header: "Notes",
        key: "notes",
        width: 50,
      },
    ];

    customers.forEach((customer) => {
      worksheet.addRow(customer);
    });

    const backupPath =
      "D:/CRM Backup";

    if (!fs.existsSync(backupPath)) {
      fs.mkdirSync(backupPath, {
        recursive: true,
      });
    }

    const fileName = `backup-${new Date()
      .toISOString()
      .split("T")[0]}.xlsx`;

    const fullPath =
      path.join(backupPath, fileName);

    await workbook.xlsx.writeFile(
      fullPath
    );

    console.log(
      "Backup created:",
      fullPath
    );
  } catch (err) {
    console.log(err);
  }
}

/* DAILY BACKUP */

cron.schedule("0 0 * * *", () => {
  createBackup();
});

/* MANUAL BACKUP */

app.get("/manual-backup", async (req, res) => {
  await createBackup();

  res.json({
    success: true,
  });
});

/* ---------------- SERVER ---------------- */

const PORT =
  process.env.PORT || 10000;

app.listen(PORT, () => {
  console.log(
    `Server running on port ${PORT}`
  );
});
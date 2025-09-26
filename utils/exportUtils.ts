import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

export type ExportFormat = "csv" | "xlsx";

// Generic export function for any data
export const exportData = async (
  data: any[],
  filename: string,
  format: ExportFormat = "xlsx",
  sheetName: string = "Data"
) => {
  if (!data || data.length === 0) {
    alert("No data to export");
    return;
  }

  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(sheetName);

    // Add headers dynamically
    if (data.length > 0) {
      worksheet.columns = Object.keys(data[0]).map((key) => ({
        header: key,
        key: key,
        width: 20,
      }));
    }

    // Add rows
    worksheet.addRows(data);

    if (format === "csv") {
      // Generate CSV
      const csvBuffer = await workbook.csv.writeBuffer();
      const blob = new Blob([csvBuffer], { type: "text/csv;charset=utf-8" });
      saveAs(blob, `${filename}.csv`);
    } else {
      // Generate XLSX
      const excelBuffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([excelBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      saveAs(blob, `${filename}.xlsx`);
    }

    console.log(`Successfully exported ${data.length} records as ${format.toUpperCase()}`);
  } catch (error) {
    console.error("Export failed:", error);
    alert("Export failed. Please try again.");
  }
};

// Specific export functions for different data types
export const exportOrders = (orders: any[], format: ExportFormat = "xlsx") => {
  const formattedData = orders.map((order) => ({
    "Order ID": order.id,
    Customer: order.customerName,
    Date: order.date,
    "Total Amount": order.total,
    Status: order.status,
    "Payment Method": order.paymentMethod || "",
    "Items Count": order.orderItems?.length || 0,
    "Assigned User": order.assignedUserId || "",
    "Cheque Balance": order.chequeBalance || 0,
    "Credit Balance": order.creditBalance || 0,
    Notes: order.notes || "",
  }));

  const timestamp = new Date().toISOString().split("T")[0];
  exportData(formattedData, `orders_${timestamp}`, format, "Orders");
};

export const exportProducts = (products: any[], format: ExportFormat = "xlsx") => {
  const formattedData = products.map((product) => ({
    "Product ID": product.id,
    Name: product.name,
    Category: product.category,
    Price: product.price,
    Stock: product.stock,
    SKU: product.sku,
    Supplier: product.supplier,
    "Image URL": product.imageUrl || "",
  }));

  const timestamp = new Date().toISOString().split("T")[0];
  exportData(formattedData, `products_${timestamp}`, format, "Products");
};

export const exportCustomers = (customers: any[], format: ExportFormat = "xlsx") => {
  const formattedData = customers.map((customer) => ({
    "Customer ID": customer.id,
    Name: customer.name,
    Email: customer.email,
    Phone: customer.phone,
    Location: customer.location,
    "Join Date": customer.joinDate,
    "Total Spent": customer.totalSpent,
    "Outstanding Balance": customer.outstandingBalance,
  }));

  const timestamp = new Date().toISOString().split("T")[0];
  exportData(formattedData, `customers_${timestamp}`, format, "Customers");
};

export const exportDriverAllocations = (allocations: any[], format: ExportFormat = "xlsx") => {
  const formattedData = allocations.map((allocation) => ({
    "Allocation ID": allocation.id,
    "Driver ID": allocation.driverId,
    "Driver Name": allocation.driverName,
    Date: allocation.date,
    "Allocated Items": JSON.stringify(allocation.allocatedItems),
    "Returned Items": JSON.stringify(allocation.returnedItems || []),
    "Sales Total": allocation.salesTotal,
    Status: allocation.status,
  }));

  const timestamp = new Date().toISOString().split("T")[0];
  exportData(formattedData, `driver_allocations_${timestamp}`, format, "Driver Allocations");
};

export const exportDriverSales = (sales: any[], format: ExportFormat = "xlsx") => {
  const formattedData = sales.map((sale) => ({
    "Sale ID": sale.id,
    "Driver ID": sale.driverId,
    "Allocation ID": sale.allocationId,
    Date: sale.date,
    "Customer Name": sale.customerName,
    "Customer ID": sale.customerId || "",
    Total: sale.total,
    "Amount Paid": sale.amountPaid,
    "Credit Amount": sale.creditAmount,
    "Payment Method": sale.paymentMethod,
    "Payment Reference": sale.paymentReference || "",
    Notes: sale.notes || "",
    "Sold Items": JSON.stringify(sale.soldItems),
  }));

  const timestamp = new Date().toISOString().split("T")[0];
  exportData(formattedData, `driver_sales_${timestamp}`, format, "Driver Sales");
};

export const exportUsers = (users: any[], format: ExportFormat = "xlsx") => {
  const formattedData = users.map((user) => ({
    "User ID": user.id,
    Name: user.name,
    Email: user.email,
    Phone: user.phone || "",
    Role: user.role,
    Status: user.status,
    "Last Login": user.lastLogin || "",
    "Assigned Suppliers": Array.isArray(user.assignedSupplierNames)
      ? user.assignedSupplierNames.join(", ")
      : "",
  }));

  const timestamp = new Date().toISOString().split("T")[0];
  exportData(formattedData, `users_${timestamp}`, format, "Users");
};

export const exportSuppliers = (suppliers: any[], format: ExportFormat = "xlsx") => {
  const formattedData = suppliers.map((supplier) => ({
    "Supplier ID": supplier.id,
    Name: supplier.name,
    "Contact Person": supplier.contactPerson,
    Email: supplier.email,
    Phone: supplier.phone,
    Address: supplier.address,
    "Join Date": supplier.joinDate,
  }));

  const timestamp = new Date().toISOString().split("T")[0];
  exportData(formattedData, `suppliers_${timestamp}`, format, "Suppliers");
};

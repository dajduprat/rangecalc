/* global clearInterval, console, setInterval */

/**
 * Builds an Excel.EntityCellValue with 10 fields and 2 nested entities
 * @customfunction
 * @returns {object} An entity cell value object with 10 properties and 2 nested entities
 */
export async function entity() {
  // Add delay of 1 second
  await new Promise((resolve) => setTimeout(resolve, 1000));

  const nestedEntity1 = {
    type: "Entity",
    text: "Address Information",
    properties: {
      street: { type: "String", basicValue: "123 Main Street" },
      city: { type: "String", basicValue: "New York" },
      zipCode: { type: "Double", basicValue: 10001 },
    },
    layouts: {
      card: {
        title: { property: "city" },
        sections: [
          {
            layout: "List",
            properties: ["street", "city", "zipCode"],
          },
        ],
      },
    },
  };

  const nestedEntity2 = {
    type: "Entity",
    text: "Contact Details",
    properties: {
      email: { type: "String", basicValue: "john.doe@example.com" },
      phone: { type: "String", basicValue: "+1-555-0123" },
      department: { type: "String", basicValue: "Engineering" },
    },
    layouts: {
      card: {
        title: { property: "department" },
        sections: [
          {
            layout: "List",
            properties: ["email", "phone", "department"],
          },
        ],
      },
    },
  };

  const entityValue = {
    type: "Entity",
    text: "Employee Record",
    properties: {
      employeeId: { type: "Double", basicValue: 12345 },
      firstName: { type: "String", basicValue: "John" },
      lastName: { type: "String", basicValue: "Doe" },
      salary: { type: "Double", basicValue: 85000, numberFormat: "$* #,##0.00" },
      hireDate: { type: "String", basicValue: "2020-03-15" },
      isActive: { type: "Boolean", basicValue: true },
      yearsOfService: { type: "Double", basicValue: 4 },
      performanceRating: { type: "Double", basicValue: 4.5 },
      projectsCompleted: { type: "Double", basicValue: 18 },
      address: nestedEntity1,
      contact: nestedEntity2,
    },
    layouts: {
      card: {
        title: { property: "firstName" },
        sections: [
          {
            layout: "List",
            properties: ["employeeId", "firstName", "lastName"],
          },
          {
            layout: "List",
            title: "Employment Info",
            collapsible: true,
            collapsed: false,
            properties: ["salary", "hireDate", "yearsOfService"],
          },
          {
            layout: "List",
            title: "Performance",
            collapsed: true,
            properties: ["performanceRating", "projectsCompleted", "isActive"],
          },
          {
            layout: "List",
            title: "Details",
            collapsed: true,
            properties: ["address", "contact"],
          },
        ],
      },
    },
  };

  return entityValue;
}

/**
 * Returns the current timestamp
 * @customfunction
 * @streaming
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @requiresAddress
 * @returns {string}
 */
export function streamingTimestamp(invocation) {
  // Send initial timestamp
  console.log("Streaming address cell value: " + invocation.address);
  invocation.setResult(new Date().toISOString());
}

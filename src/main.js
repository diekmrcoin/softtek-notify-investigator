import prompt from "prompt";
import { promisify } from "util";
import { WorkbookManager } from "./xlsx.js";
const get = promisify(prompt.get);
// Start the prompt
prompt.start();

// Define the prompt schema
const mainSchema = {
  properties: {
    menu: {
      description:
        "Choose an option: 1) Set date, 2) Populate Excel, 3) Bundle into zip, 4) Send email, 5) Exit",
      pattern: /^[1-5]$/,
      message: "Option must be between 1 and 5",
      required: true,
      before: (value) => +value,
    },
  },
};

// Get the user's input
get(mainSchema).then(async ({ menu }) => {
  switch (menu) {
    case 1:
      console.log("You chose to set date");
      // Do something to set date
      break;
    case 2:
      console.log("You chose to populate Excel");
      const dateSchema = {
        properties: {
          year: {
            description: "Enter the year",
            pattern: /^[0-9]{4}$/,
            message: "Year must be a four-digit number",
            required: true,
            before: (value) => +value,
          },
          month: {
            description: "Enter the month",
            pattern: /^(1[0-2]|0?[1-9])$/,
            message: "Month must be a number between 1 and 12",
            required: true,
            before: (value) => +value,
          },
          day: {
            description: "Enter the day",
            pattern: /^(3[01]|[12][0-9]|0?[1-9])$/,
            message: "Day must be a number between 1 and 31",
            required: true,
            before: (value) => +value,
          },
        },
      };
      let { year, month, day } = await get(dateSchema);
      const manager = new WorkbookManager(year, month, day);
      const workbook = manager.setup();
      manager.populate(workbook);
      break;
    case 3:
      console.log("You chose to bundle into zip");
      // Do something to bundle into zip
      break;
    case 4:
      console.log("You chose to send email");
      // Do something to send email
      break;
    default:
      console.log("You chose to exit");
      // Exit the application
      process.exit();
  }
});

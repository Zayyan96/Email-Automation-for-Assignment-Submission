function sendEmails() {
  var sheetId1 = "12Cma0qxmsz89EsaK2U6wALvPud2QLwZHGx0UzmS2fIM"; // ID of the first Google Sheet
  var sheet1 = SpreadsheetApp.openById(sheetId1);
  if (!sheet1) {
    Logger.log("Sheet not found with ID: " + sheetId1);
    return;
  }
  var values = sheet1.getDataRange().getValues();
  var lastProcessedRow = getLastProcessedRow(); // Retrieve the last processed row
  var sheetId2 = "1PiUJ1cO1U-YMfVC4KNQIShmekp7wM3rR6JagQHG0FP0"; // ID of the second Google Sheet
  var sheet2 = SpreadsheetApp.openById(sheetId2);
  if (!sheet2) {
    Logger.log("Sheet not found with ID: " + sheetId2);
    return;
  }
  var data = sheet2.getDataRange().getValues();
  Logger.log(lastProcessedRow);
  var subject1;
  // Iterate through each row starting from the second row (assuming the first row contains headers)
  for (var i = lastProcessedRow; i < values.length; i++) {
    var studentId = values[i][2];
    var emailAddress = null;
    var emailAddress1 = null;
    var emailAddress2 = null;
    // Iterate through the second sheet to find the matching student ID
    for (var j = 1; j < data.length; j++) {
      if (studentId == data[j][0]) {
        emailAddress = data[j][1];
        emailAddress1 = data[j][4];
        emailAddress2 = data[j][5];
        Logger.log("Email Address found for Student ID: " + studentId);
        break; // Exit the loop once a match is found
      }
    }

    // Check if email address is found
    if (emailAddress) {
      // Customize the subject and body of your email based on student's data
      var subject = "Confirmation of Successful Assignment Submission by " + values[i][1];
      var htmlBodyContent =
        "<p>Dear " + values[i][1] + ",</p>" +
        "<p>We trust this email finds you well. We would like to acknowledge successful submission of your recent assignment.</p>" +
        "<p>Here are the details confirming your assignment submission:</p>" +
        "<ul>" +
        "<li><strong>Student ID:</strong> " + values[i][2] + "</li>" +
        "<li><strong>Name:</strong> " + values[i][1] + "</li>" +
        "<li><strong>Grade:</strong> " + (values[i][3] || "Not specified") + "</li>" +
        "<li><strong>Assignment's Subject:</strong> " + (values[i][4] || values[i][13] || values[i][22] || values[i][31]) + "</li>" +
        "<li><strong>Assignment Name:</strong> " + (values[i][5] || values[i][7] || values[i][9] || values[i][11] || values[i][14] || values[i][16] || values[i][18] || values[i][20] || values[i][23] || values[i][25] || values[i][27] || values[i][29] || values[i][32] || values[i][34] || values[i][36] || values[i][38]) + "</li>" +
        "</ul>" +
        "<p>We appreciate your dedication to meeting deadlines and adhering to the guidelines provided. The successful completion of your assignment is a testament to your hard work and understanding of the course material.</p>" +
        "<p>If you have any concerns or questions regarding the submission process or if you require further clarification on the assignment feedback, please do not hesitate to reach out to your course instructor or the designated contact person for your program.</p>" +
        "<p>Once again, congratulations on completing and submitting your assignment. We look forward to your continued success in your academic endeavors.</p>" +
        "<p>Best regards,<br/>Management Online Tuitions</p>";

      var htmlBodyContent1 =
        "<p>Dear Parents</p>" +
        "<p>We trust this email finds you well. We would like to acknowledge successful submission of your recent assignment.</p>" +
        "<p>Here are the details confirming your assignment submission:</p>" +
        "<ul>" +
        "<li><strong>Student ID:</strong> " + values[i][2] + "</li>" +
        "<li><strong>Name:</strong> " + values[i][1] + "</li>" +
        "<li><strong>Grade:</strong> " + (values[i][3] || "Not specified") + "</li>" +
        "<li><strong>Assignment's Subject:</strong> " + (values[i][4] || values[i][13] || values[i][22] || values[i][31]) + "</li>" +
        "<li><strong>Assignment Name:</strong> " + (values[i][5] || values[i][7] || values[i][9] || values[i][11] || values[i][14] || values[i][16] || values[i][18] || values[i][20] || values[i][23] || values[i][25] || values[i][27] || values[i][29] || values[i][32] || values[i][34] || values[i][36] || values[i][38]) + "</li>" +
        "</ul>" +
        "<p>We appreciate your dedication to meeting deadlines and adhering to the guidelines provided. The successful completion of your assignment is a testament to your hard work and understanding of the course material.</p>" +
        "<p>If you have any concerns or questions regarding the submission process or if you require further clarification on the assignment feedback, please do not hesitate to reach out to your course instructor or the designated contact person for your program.</p>" +
        "<p>Once again, congratulations on completing and submitting your assignment. We look forward to your continued success in your academic endeavors.</p>" +
        "<p>Best regards,<br/>Management Online Tuitions</p>";
      subject1 = values[i][4] || values[i][13] || values[i][22] || values[i][31];
      Logger.log(subject1);
      try {
        GmailApp.sendEmail(emailAddress, subject, null, {
          htmlBody: htmlBodyContent,
          from: "info@online-tuitions.com"
        });
        // if (emailAddress1) {                                                       DUE TO EMAIL QUOTA RESTRICTION
        //   GmailApp.sendEmail(emailAddress1, subject, null, {
        //     htmlBody: htmlBodyContent1,
        //     from: "info@online-tuitions.com"
        //   });
        // }
        // if (emailAddress2) {
        //   GmailApp.sendEmail(emailAddress2, subject, null, {
        //     htmlBody: htmlBodyContent1,
        //     from: "info@online-tuitions.com"
        //   });
        // }
        // if (subject1 == "Chemistry") {
        //   // Send email to the Chemistry teacher
        //   GmailApp.sendEmail("hamzatariq017@gmail.com", subject, null, {
        //     htmlBody: htmlBodyContent,
        //     from: "info@online-tuitions.com"
        //   });
        // } else if (subject1 == "Physics" || subject1 == "Mathematics") {
        //   // Send email to the Physics and Maths teacher
        //   GmailApp.sendEmail("hassantariq539@gmail.com", subject, null, {
        //     htmlBody: htmlBodyContent,
        //     from: "info@online-tuitions.com"
        //   });
        // }
        updateLastProcessedRow(values.length);
      } catch (error) {
        var senderEmail = Session.getActiveUser().getEmail();
        var errorSubject = "Error: Failed to Send Email to " + error.toString();
        var errorMessage = "Email sending failed to the recipient with this wrong email: " + error.toString() + "<p>Error details: </p>" +
          "<ul>" +
          "<li><strong>Student ID:</strong> " + values[i][2] + "</li>" +
          "<li><strong>Name:</strong> " + values[i][1] + "</li>" +
          "<li><strong>Grade:</strong> " + (values[i][3] || "Not specified") + "</li>" +
          "<li><strong>Assignment's Subject:</strong> " + (values[i][4] || values[i][13] || values[i][22] || values[i][31]) + "</li>" +
          "<li><strong>Assignment Name:</strong> " + (values[i][5] || values[i][7] || values[i][9] || values[i][11] || values[i][14] || values[i][16] || values[i][18] || values[i][20] || values[i][23] || values[i][25] || values[i][27] || values[i][29] || values[i][32] || values[i][34] || values[i][36] || values[i][38]) + "</li>" +
          "</ul>";

        GmailApp.sendEmail(senderEmail, errorSubject, null, {
          htmlBody: errorMessage, // corrected property name
          from: "info@online-tuitions.com"
        });
        continue; // Move to the next iteration without updating lastProcessedRow
      }
    } else {
      // If email address is not found, send an email to the sender
      var senderEmail = Session.getActiveUser().getEmail();
      var notFoundSubject = "Error: Email Address Not Found for Student ID " + studentId;
      var notFoundMessage = "Email address not found for Student ID: " + studentId;
      GmailApp.sendEmail(senderEmail, notFoundSubject, notFoundMessage, {
        from: "info@online-tuitions.com"
      });
    }

    // Add a delay to avoid reaching email sending quotas
    Utilities.sleep(5000); // Sleep for 1 second (you can adjust this as needed)
    updateLastProcessedRow(values.length);
  }
}
// Helper function to get the last processed row
function getLastProcessedRow() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastProcessedRow = scriptProperties.getProperty('lastProcessedRow');
  if (lastProcessedRow) {
    return parseInt(lastProcessedRow);
  } else {
    return 0; // Start from the first row if no last processed row is found
  }
}

// Helper function to update the last processed row
function updateLastProcessedRow(row) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('lastProcessedRow', row.toString());
}

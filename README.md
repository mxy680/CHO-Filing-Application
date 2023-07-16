# Cohens-Fashion-Optical-Filing-Application
I developed this application to streamline the process of filing patients' documents for Cohens-Fashion-Optical, an optical retailer. The previous manual process was tedious and time-consuming. With the automation provided by my script, employees can now easily input scanned batches of documents with minimal effort.

The script utilizes Optical Character Recognition (OCR) capabilities from the pytesseract library along with regular expressions (regex) to accurately extract relevant patient information from the forms. To ensure stability, I implemented default null values for patient information in cases where OCR errors may occur.

Selenium, a web automation tool, is used to navigate through the retailer's website and input the extracted patient information. Although the website was not designed for automation, I implemented various workarounds to handle any potential issues. For example, I incorporated time delays using the time.sleep() function to allow for proper page loading. Additionally, I implemented the ensure_click() function to handle scenarios where an element may not be immediately clickable.

In cases where a patient could not be found, the script records their information and location within the batch in a dedicated CSV file. This allows users to easily locate the patient's file and manually input their document, ensuring no information is lost.

The script is designed to run seamlessly without interruptions. Users simply need to specify the batch number and form type, and then initiate the program by clicking the "run" button. The automation process eliminates human errors and greatly enhances efficiency.

Overall, this application significantly improves the document filing process for Cohens-Fashion-Optical, reducing manual effort and enhancing accuracy.

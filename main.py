import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By

# "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\EdgeUserData"

class WebsiteTextExporter:
    def __init__(self, port=9222):
        edge_options = Options()
        edge_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
        self.driver = webdriver.Edge(options=edge_options)

    def extract_texts(self):
        """Extract all visible texts from the webpage"""
        elements = self.driver.find_elements(By.XPATH, "//*[not(*) and normalize-space()]")
        texts = list(set([element.text.strip() for element in elements if element.text.strip()]))
        return texts

    def delete_texts(self, texts):
        """Delete specified texts from the webpage"""
        for text in texts:
            try:
                # Handle different quote types in XPath
                if '"' in text:
                    escaped_text = text.replace("'", "''")
                    xpath = f"//*[not(*) and normalize-space() = '{escaped_text}']"
                else:
                    xpath = f'//*[not(*) and normalize-space() = "{text}"]'

                elements = self.driver.find_elements(By.XPATH, xpath)
                for element in elements:
                    # Remove text content completely
                    self.driver.execute_script("arguments[0].textContent = '';", element)
            except Exception as e:
                print(f"Error deleting text '{text}': {e}")

    def export_to_excel(self, filename="website_texts.xlsx"):
        """Export extracted texts to Excel and delete them from the page"""
        texts = self.extract_texts()
        if not texts:
            print("No texts found to export.")
            return

        # Save to Excel
        df = pd.DataFrame(texts, columns=["Exported Texts"])
        df.to_excel(filename, index=False)
        print(f"Exported {len(texts)} texts to {filename}")

        # Delete texts from webpage
        self.delete_texts(texts)
        print("Deleted exported texts from webpage. Remaining texts were not captured.")

    def close(self):
        """Close the browser"""
        self.driver.quit()


# Usage
if __name__ == "__main__":
    exporter = WebsiteTextExporter()
    exporter.export_to_excel()
    exporter.close()

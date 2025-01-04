def send_file(driver, phone_number, file_path, product):
    try:
        if not os.path.exists(file_path):
            update_log(f"File does not exist: {file_path}")
            return

        # Open chat with the phone number
        driver.get(f"https://web.whatsapp.com/send?phone={phone_number}")
        update_log(f"Opened chat with {phone_number}")

        # Wait for and click the "Attach" button
        attach_btn = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[1]/div[2]/button'))
        )
        attach_btn.click()

        # Wait for the file input and upload the file
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@type="file"]'))
        )
        file_input.send_keys(file_path)
        update_log(f"Uploaded file '{os.path.basename(file_path)}' for {phone_number}")

        # Wait for the "Send" button and click
        send_btn = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/div[3]/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div'))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", send_btn)
        send_btn.click()

        time.sleep(4)  # Ensure the file is sent
        update_log(f"File '{os.path.basename(file_path)}' for '{product}' sent to {phone_number}")

    except Exception as e:
        update_log(f"Failed to send file '{os.path.basename(file_path)}' to {phone_number} - {str(e)}")def send_file(driver, phone_number, file_path, product):
    try:
        if not os.path.exists(file_path):
            update_log(f"File does not exist: {file_path}")
            return

        # Open chat with the phone number
        driver.get(f"https://web.whatsapp.com/send?phone={phone_number}")
        update_log(f"Opened chat with {phone_number}")

        # Wait for and click the "Attach" button
        attach_btn = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//div[@title="Attach"]'))
        )
        attach_btn.click()

        # Wait for the file input and upload the file
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@type="file"]'))
        )
        file_input.send_keys(file_path)
        update_log(f"Uploaded file '{os.path.basename(file_path)}' for {phone_number}")

        # Wait for the "Send" button and click
        send_btn = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//span[@data-icon="send"]'))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", send_btn)
        send_btn.click()

        time.sleep(4)  # Ensure the file is sent
        update_log(f"File '{os.path.basename(file_path)}' for '{product}' sent to {phone_number}")

    except Exception as e:
        update_log(f"Failed to send file '{os.path.basename(file_path)}' to {phone_number} - {str(e)}")
# Function to fetch images based on the document list
def fetch_images():
    # Get document file path
    doc_file = filedialog.askopenfilename(filetypes=[("All Files", "*.xlsx *.xls *.csv *.txt")])
    if not doc_file:
        return

    # Read the document based on its extension
    file_extension = os.path.splitext(doc_file)[1].lower()

    try:
        if file_extension in ['.xlsx', '.xls']:
            df = pd.read_excel(doc_file)
        elif file_extension == '.csv':
            df = pd.read_csv(doc_file)
        elif file_extension == '.txt':
            df = pd.read_csv(doc_file, delimiter="\t", header=None, names=["prd_code"])
        else:
            messagebox.showerror("Error", "Unsupported file format.")
            return
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read the document: {str(e)}")
        return

    # Column name for product codes (adjust accordingly)
    image_column = 'prd_code'

    if image_column not in df.columns:
        messagebox.showerror("Error", f"'{image_column}' column not found in the document.")
        return

    # Get image directory
    image_directory = filedialog.askdirectory(title="Select Image Directory")
    if not image_directory:
        return

    # Get output directory
    output_directory = filedialog.askdirectory(title="Select Output Directory")
    if not output_directory:
        return

    # Normalize paths
    image_directory = os.path.normpath(image_directory)
    output_directory = os.path.normpath(output_directory)

    # Create output directory if it doesn't exist
    os.makedirs(output_directory, exist_ok=True)

    # Get all images in the directory and subdirectories
    available_images = find_images_recursively(image_directory)
    print(f"Available images: {available_images}")

    # Get images that already exist in the output directory
    existing_images = get_existing_images(output_directory)
    print(f"Existing images in output directory: {existing_images}")

    # Track if any image is successfully copied
    copied_images = 0
    not_found_images = []  # List to track images not found

    # Loop through the document list and try to find matching images
    for index, row in df.iterrows():
        image_name = normalize_image_name(str(row[image_column]).strip())  # Normalize document image name

        # Check if the image is already in the output directory
        if image_name in existing_images:
            print(f"Image already exists in output directory: {image_name}")
            continue  # Skip this image since it's already in the output folder

        # Check if the normalized image name is a substring of any available image name
        found_image_path = None
        for available_name, image_path in available_images.items():
            if image_name in available_name:  # Check for substring match
                found_image_path = image_path
                break

        if found_image_path:
            try:
                # Read and save the image
                image = cv2.imread(found_image_path)
                if image is not None:
                    output_image_path = os.path.join(output_directory, os.path.basename(found_image_path))
                    cv2.imwrite(output_image_path, image)
                    copied_images += 1
                    print(f"Found and saved: {found_image_path}")
                else:
                    print(f"Failed to read image: {found_image_path}")
            except Exception as e:
                print(f"Error processing image {found_image_path}: {e}")
        else:
            print(f"Image not found: {row[image_column]}")
            not_found_images.append(row[image_column])  # Add to not-found list

    # Save the not-found images list to a file
    if not_found_images:
        not_found_file = os.path.join(output_directory, "not_found_images.csv")
        pd.DataFrame(not_found_images, columns=[image_column]).to_csv(not_found_file, index=False)
        print(f"Not found images saved to: {not_found_file}")

    if copied_images > 0:
        messagebox.showinfo("Success", f"Successfully fetched and saved {copied_images} images!")
    else:
        messagebox.showwarning("No Images Found", "No new images were found or copied.")
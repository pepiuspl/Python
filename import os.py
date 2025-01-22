import os
import shutil

def copy_cooperation_terms_for_all_clients(base_path):
    # Paths for 2024 and 2025 directories
    year_2024 = os.path.join(base_path, "_____2024")
    year_2025 = os.path.join(base_path, "______2025")

    # Path to the fallback PDF on the desktop (updated with the correct file name)
    fallback_pdf = r"C:\Users\tomasz.plewka\Desktop\ZASADY_BAZOWE.pdf"

    # Ensure both 2024 and 2025 directories exist
    if not os.path.exists(year_2024):
        print(f"The 2024 folder does not exist at: {year_2024}")
        return

    if not os.path.exists(year_2025):
        os.makedirs(year_2025)  # Create 2025 folder if it doesn't exist

    # Iterate through all regions in 2024
    for region in os.listdir(year_2024):
        region_2024_path = os.path.join(year_2024, region)
        region_2025_path = os.path.join(year_2025, region)

        # Check if the region folder is a directory
        if os.path.isdir(region_2024_path):
            # Ensure the region folder exists in 2025
            if not os.path.exists(region_2025_path):
                os.makedirs(region_2025_path)

            # Iterate through all clients in the region
            for client in os.listdir(region_2024_path):
                client_2024_path = os.path.join(region_2024_path, client)
                client_2025_path = os.path.join(region_2025_path, client)

                # Check if the client folder is a directory
                if os.path.isdir(client_2024_path):
                    # Ensure the client folder exists in 2025
                    if not os.path.exists(client_2025_path):
                        os.makedirs(client_2025_path)

                    # Look for folders containing "ZASADY" in their name in 2024
                    cooperation_terms_2024 = None
                    for folder in os.listdir(client_2024_path):
                        folder_path = os.path.join(client_2024_path, folder)
                        if os.path.isdir(folder_path) and "ZASADY" in folder:
                            cooperation_terms_2024 = folder_path
                            break  # We found the folder, no need to check others

                    if cooperation_terms_2024:
                        # Path for the folder in 2025
                        cooperation_terms_2025 = os.path.join(client_2025_path, os.path.basename(cooperation_terms_2024))

                        # Check if the folder exists in 2025
                        if os.path.exists(cooperation_terms_2025):
                            # Copy files from 2024 to 2025
                            try:
                                for file_name in os.listdir(cooperation_terms_2024):
                                    file_2024_path = os.path.join(cooperation_terms_2024, file_name)
                                    file_2025_path = os.path.join(cooperation_terms_2025, file_name)

                                    # Only copy files (skip directories)
                                    if os.path.isfile(file_2024_path):
                                        shutil.copy(file_2024_path, file_2025_path)
                                        print(f"Copied file '{file_name}' for client: {client} in region: {region}")
                                    else:
                                        print(f"Skipping directory '{file_name}' (only files will be copied)")
                            except Exception as e:
                                print(f"Error copying files from 'ZASADY' folder for client: {client} in region: {region}. Error: {e}")
                        else:
                            print(f"Folder 'ZASADY' does not exist for client: {client} in region: {region} in 2025")
                    else:
                        # If no "ZASADY" folder found in 2024, copy the fallback PDF to 2025
                        print(f"No 'ZASADY' folder found for client: {client} in region: {region}. Copying fallback PDF.")
                        # Ensure the target folder exists in 2025
                        cooperation_terms_2025 = os.path.join(client_2025_path, "9_ZASADY WSPÓŁPRACY")
                        if not os.path.exists(cooperation_terms_2025):
                            os.makedirs(cooperation_terms_2025)

                        # Copy the fallback PDF to the 2025 folder
                        try:
                            shutil.copy(fallback_pdf, os.path.join(cooperation_terms_2025, os.path.basename(fallback_pdf)))
                            print(f"Copied fallback PDF for client: {client} in region: {region}")
                        except Exception as e:
                            print(f"Error copying fallback PDF for client: {client} in region: {region}. Error: {e}")
                else:
                    print(f"'{client}' is not a valid client directory in region: {region}")
        else:
            print(f"'{region}' is not a valid region directory in 2024")

# Replace this with the path to your "TECZKI KLIENTÓW" folder
base_path = r"M:\TECZKI KLIENTÓW"
copy_cooperation_terms_for_all_clients(base_path)

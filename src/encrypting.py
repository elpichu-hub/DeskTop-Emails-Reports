import pandas as pd
import base64
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad, unpad
import os
import smart_sheet_poster

def main_func_encrypt(file_path, key):
    def encrypt_data(file_path, key):
        df = pd.read_csv(file_path)

        # Encrypt each cell in the DataFrame
        encrypted_data = []
        for index, row in df.iterrows():
            encrypted_row = {}
            for column in df.columns:
                value = str(row[column])
                iv, encrypted = encrypt_message(value, key)
                encrypted_row[column] = f"{iv}:{encrypted}"
            encrypted_data.append(encrypted_row)

        # Convert the encrypted data to a DataFrame and save it
        encrypted_df = pd.DataFrame(encrypted_data)
        encrypted_file_path = 'encrypted_lookups.csv'
        encrypted_df.to_csv(encrypted_file_path, index=False)

        os.startfile(encrypted_file_path)
        return encrypted_file_path

    def encrypt_message(message, key):
        cipher = AES.new(key, AES.MODE_CBC)
        padded_message = pad(message.encode(), AES.block_size)
        encrypted_message = cipher.encrypt(padded_message)
        return base64.b64encode(cipher.iv).decode(), base64.b64encode(encrypted_message).decode()

    
    encrypted_file_path = encrypt_data(file_path, key)

    return encrypted_file_path
    


def main_func_decrypt(file_path, key):
    def decrypt_data(file_path, key):
        encrypted_df = pd.read_csv(file_path)

        # Decrypt each cell in the DataFrame
        decrypted_data = []
        for index, row in encrypted_df.iterrows():
            decrypted_row = {}
            for column in encrypted_df.columns:
                iv_b64, encrypted_b64 = row[column].split(':')
                decrypted_row[column] = decrypt_message(iv_b64, encrypted_b64, key)
            decrypted_data.append(decrypted_row)

        # Convert the decrypted data to a DataFrame and display it
        decrypted_df = pd.DataFrame(decrypted_data)
        decrypted_df.to_csv('decrypted_lookups.csv', index=False)
        print(decrypted_df.head())

    def decrypt_message(iv_b64, encrypted_b64, key):
        iv = base64.b64decode(iv_b64)
        encrypted_message = base64.b64decode(encrypted_b64)
        cipher = AES.new(key, AES.MODE_CBC, iv)
        decrypted_message = cipher.decrypt(encrypted_message)
        return unpad(decrypted_message, AES.block_size).decode()
    
    decrypt_data(file_path, key)

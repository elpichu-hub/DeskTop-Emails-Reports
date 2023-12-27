import socket

# Create a TCP/IP socket
target_host = "127.0.0.1"  # Use the loopback address for local testing
target_port = 9998

# Connect the socket to the server's IP and port
client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
client.connect((target_host, target_port))

# Send some data to the server
request = "GET / HTTP/1.1\r\nHost: google.com\r\n\r\n"
client.send(request.encode())

# Receive the server's response
response = client.recv(4096)
print(response.decode())

# Close the client socket
client.close()

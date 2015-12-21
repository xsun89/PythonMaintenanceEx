import socket;
import time;

sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM);
print("正在和Jinko连线...");
sock.connect(('localhost', 8005));
print("");


while True:
	speak = input("和Jinko说点什么:");

	if speak == "quit":
		break;

	if speak == "":
		continue;

	# print("发送中..." + "("+ speak +")")
	sock.send(speak.encode('utf-8'));
	# print("已发送")

	print("Jinko在思考...");
	answer = sock.recv(8192);
	print("Jinko回复你:" + answer.decode('utf-8'));
	print("");

sock.close();
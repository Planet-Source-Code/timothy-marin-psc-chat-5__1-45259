//By Timothy Marin
//Copyright 2003
//www.IntraDream.com
//See Example At [ idream.noip.com:56/Client.swf ]
setInterval( sendPing, 30000 );
var LastOb;
var Linfo;
setInterval( sendType, 10 );
function connect() {
	mySocket = new XMLSocket();
	mySocket.onConnect = handleConnect;
	mySocket.onClose = handleClose;
	mySocket.onData = handleIncoming;
	mySocket.connect(null, 15003); //Can Only Connect to Sending server(null) when coming from web
	mySocket.host = host;
	mySocket.port = port;
}
function handleConnect(succeeded) {
	if (succeeded) {
		mySocket.connected = true;
		mySocket.send(LINFO);
		gotoAndStop(42);
		Selection.setFocus("_level0.outgoing");
	} else {
		gotoAndStop(29);
	}
}
function handleClose() {
	incoming += ("Disconnected From Server, Refresh the Page to Reconnect...."+newline);
	mySocket.connected = false;
}
function handleIncoming(messageObj) {
		$DataArrayA = messageObj.toString().split(chr(1));
		if ($DataArrayA[0] == "join") {
			$DataArrayB = $DataArrayA[1].split(chr(2));
			list.addItem($DataArrayb[0],$DataArrayb[1]);
			list1.addItem($DataArrayb[0]);
			mySocket.send("flash"+chr(1));
			incoming += ("* "+$DataArrayb[0]+" Has Joined..."+newline);
		} else if ($DataArrayA[0] == "status") {
			$DataArrayB = $DataArrayA[1].split(chr(2));
			$GumEntries = list.getLength();
			for ($z=0; $z<$GumEntries; $z++) {
				if (list.getItemAt($z).label == $DataArrayb[0]) {
					list.removeItemAt($z)
					list.addItem($DataArrayb[0],$DataArrayb[1]);
					list1.removeItemAt(number($z)+1)
					list1.addItem($DataArrayb[0]);
				}
			}
		} else if ($DataArrayA[0] == "topic") {
			Topic=$DataArrayA[1]
		} else if ($DataArrayA[0] == "server") {
			incoming += (" Server : "+$DataArrayA[1]+newline);
		} else if ($DataArrayA[0] == "part") {
			$DataArrayB = $DataArrayA[1].split(chr(2));
			$GumEntries = list.getLength();
			for ($z=0; $z<$GumEntries; $z++) {
				if (list.getItemAt($z).label == $DataArrayb[0]) {
					list.removeItemAt($z)
					list1.removeItemAt(number($z)+1)
				}
			}
			incoming += ("* "+$DataArrayb[0]+" Has Left..."+newline);
		} else if ($DataArrayA[0] == "list") {
			$DataArrayB = $DataArrayA[1].split(chr(2));
			list.addItem($DataArrayb[0], $DataArrayb[1]);
			list1.addItem($DataArrayb[0]);
		} else if ($DataArrayA[0] == "pmsg") {
			$DataArrayB = $DataArrayA[1].split(chr(2));
			$NumEntries = $DataArrayB.length;
				if (number($NumEntries)<9) {
					incoming += ("* "+$DataArrayb[7]+" whispers : "+$DataArrayb[6]+newline);
				} else {
					incoming += ("* You tell "+$DataArrayb[8]+" : "+$DataArrayb[6]+newline);
				}
		} else if ($DataArrayA[0] == "msg") {
			$DataArrayB = $DataArrayA[1].split(chr(2));	
			incoming += (" <"+$DataArrayb[7]+"> "+$DataArrayb[6]+newline);
		} else if ($DataArrayA[0] == "auth") {
			incoming += (" *Error Loging in or creating account."+$DataArrayA[1]+newline);
			//mysocket.handleClose
		} else {
			//Unknown MSG
		}
	incoming.scroll = incoming.maxscroll;
}
function sendType() {
	if (LastOb==outgoing) {
		//do nothing
	} else if (mySocket && mySocket.connected) {
		LastOb=outgoing;
		if (list1.getSelectedItem().label == "Main Room") {
			mySocket.send(chr(5));
		} else { 
			mySocket.send(chr(7)+chr(1)+list1.getSelectedItem().label);
		}
	}
}
function sendMessage() {
	var message = (outgoing);
	if (mySocket && mySocket.connected) {
		if (list1.getSelectedItem().label == "Main Room") {
			mySocket.send("msg"+chr(1)+"MS Sans Serif"+chr(2)+"8.25"+chr(2)+"False"+chr(2)+"False"+chr(2)+"False"+chr(2)+"#000000"+chr(2)+message);
		
		} else {
			mySocket.send("pmsg"+chr(1)+"MS Sans Serif"+chr(2)+"8.25"+chr(2)+"False"+chr(2)+"False"+chr(2)+"False"+chr(2)+"#000000"+chr(2)+message+chr(2)+list1.getSelectedItem().label);
		}
		outgoing = "";
	} else {
		incoming += "Disconnected From Server, Refresh the Page to Reconnect...."+Newline;
		incoming.scroll = incoming.maxscroll;
	}
}
function sendPing() {
	if (mySocket && mySocket.connected) {
		mySocket.send(chr(1));
	} else {
		incoming.scroll = incoming.maxscroll;
	}
}
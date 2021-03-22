


def rtpTest = "Apr:Sep"
def monthList = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
def rtpList = rtpTest.split(":")
//def periodsToPush = cscParams(monthList[monthList.indexOf(rtpList[0])..monthList.indexOf(rtpList[1])])
def periodsToPush = monthList[monthList.indexOf(rtpList[0])..monthList.indexOf(rtpList[1])]
println periodsToPush


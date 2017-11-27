package tests

import (
	"testing"
	"github.com/go-ole1c"
	"fmt"
)

func TestConnectionConstant(t *testing.T) {
	conString := "Srvr=192.168.144.1;Ref=proba;Usr=123asdf456;Pwd=winston"
	connector, _ := ole1c.CreateConnector()
	defer ole1c.DestroyConnector()
	connection, err := connector.Connect(conString)
	if err != nil {
		panic(err)
	}
	region := connection.Const("НаименованиеРЭС")
	fmt.Println("regionName: ", region.Get())
	return
}

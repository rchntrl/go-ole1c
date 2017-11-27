package ole1c

import (
	"testing"
	"fmt"
	"os"
)

type pConfiguration struct {
	Name string `db1c:"Name"`
	Version string `db1c:"Version"`
	Vendor string `db1c:"Vendor"`
}

func TestConnection(t *testing.T) {
	conString := os.Getenv("CONNECTION_STRING_1C")
	connector, _ := CreateConnector()
	defer DestroyConnector()
	connection, err := connector.Connect(conString)
	if err != nil {
		panic(err)
		return
	}
	c := new(pConfiguration )
	connection.Property("Metadata").Dispatched().Unmarshall(c)
	fmt.Printf("Наименование: %v %v \n", c.Name, c.Version)
	fmt.Printf("Поставщик: %v \n", c.Vendor)
	return
}

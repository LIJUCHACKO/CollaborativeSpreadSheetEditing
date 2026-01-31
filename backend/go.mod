module shared-spreadsheet

go 1.24.0

toolchain go1.24.11

replace python-libs => ../python-libs

require (
	github.com/gorilla/websocket v1.5.3
	github.com/kluctl/go-embed-python v0.0.0-3.13.1-20241219-1
	github.com/xuri/excelize/v2 v2.8.1
	golang.org/x/crypto v0.46.0
	python-libs v0.0.0-00010101000000-000000000000
)

require (
	github.com/gofrs/flock v0.12.1 // indirect
	github.com/mohae/deepcopy v0.0.0-20170929034955-c48cc78d4826 // indirect
	github.com/richardlehane/mscfb v1.0.4 // indirect
	github.com/richardlehane/msoleps v1.0.3 // indirect
	github.com/sirupsen/logrus v1.9.3 // indirect
	github.com/xuri/efp v0.0.0-20231025114914-d1ff6096ae53 // indirect
	github.com/xuri/nfp v0.0.0-20230919160717-d98342af3f05 // indirect
	golang.org/x/net v0.47.0 // indirect
	golang.org/x/sync v0.19.0 // indirect
	golang.org/x/sys v0.39.0 // indirect
	golang.org/x/text v0.32.0 // indirect
)

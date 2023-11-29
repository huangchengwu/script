package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"os"

	"github.com/fumiama/go-docx"
)

//发送消息返回消息id +1就是你的回复

func Dootask_sendtext(token string, payload map[string]string) (dialog_id float64, err error) {
	url := "https://t.hitosea.com/api/dialog/msg/sendtext"

	jsonPayload, _ := json.Marshal(payload)
	requestBody := bytes.NewBuffer(jsonPayload)
	log.Println("发送消息", token, string(jsonPayload))

	req, err := http.NewRequest("POST", url, requestBody)
	req.Header.Set("Content-Type", "application/json")

	req.Header.Set("Token", token)

	client := http.DefaultClient
	resp, err := client.Do(req)
	if err != nil {
		fmt.Println("Error sending HTTP request:", err)
		return
	}
	defer resp.Body.Close()

	if err != nil {
		fmt.Println("Error sending HTTP request:", err)
		return 0, err
	}
	defer resp.Body.Close()

	body, err := ioutil.ReadAll(resp.Body)

	if err != nil {
		fmt.Println("Error reading HTTP response:", err)
		return 0, err

	}

	resMap := make(map[string]interface{})
	err = json.Unmarshal(body, &resMap)
	if err != nil {
		return 0, err
	}

	dialog_id = resMap["data"].(map[string]interface{})["id"].(float64)

	dialog_id += 1

	return dialog_id, nil

}

// 获取单会话消息
func Dootask_one(token string, payload map[string]float64) (res_to string, err error) {
	url := "https://t.hitosea.com/api/dialog/one"
	jsonPayload, _ := json.Marshal(payload)
	requestBody := bytes.NewBuffer(jsonPayload)
	log.Println("发送消息", token, string(jsonPayload))

	req, err := http.NewRequest("POST", url, requestBody)
	req.Header.Set("Content-Type", "application/json")

	req.Header.Set("Token", token)

	client := http.DefaultClient
	resp, err := client.Do(req)
	if err != nil {
		fmt.Println("Error sending HTTP request:", err)
		return
	}
	defer resp.Body.Close()

	if err != nil {
		fmt.Println("Error sending HTTP request:", err)
		return "", err
	}
	defer resp.Body.Close()

	body, err := ioutil.ReadAll(resp.Body)

	if err != nil {
		fmt.Println("Error reading HTTP response:", err)
		return "", err

	}

	resMap := make(map[string]interface{})
	err = json.Unmarshal(body, &resMap)
	if err != nil {
		return "", err
	}
	log.Println(err, resMap)

	res_to = ""
	// fmt.Sprintln(resMap["data"].(map[string]interface{})["last_msg"].(map[string]interface{})["msg"].(map[string]interface{})["text"])

	return res_to, nil
}

// 登陆
func Dootask_login(payload map[string]string) (res_to string, err error) {
	url := "https://t.hitosea.com/api/users/login"

	jsonPayload, _ := json.Marshal(payload)
	requestBody := bytes.NewBuffer(jsonPayload)

	resp, err := http.Post(url, "application/json", requestBody)

	if err != nil {
		fmt.Println("Error sending HTTP request:", err)
		return "", err
	}
	defer resp.Body.Close()

	body, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		fmt.Println("Error reading HTTP response:", err)
		return "", err

	}

	resMap := make(map[string]interface{})
	err = json.Unmarshal(body, &resMap)
	if err != nil {
		return "", err
	}
	res_to = fmt.Sprintln(err, resMap["data"].(map[string]interface{})["token"])

	return res_to, nil
}

// 转换md
func convert_Md(filename, content string) {
	log.Println("md转换")
	err := ioutil.WriteFile(filename, []byte(content), 0644)
	if err != nil {
		panic(err)
	}

	println("Markdown file written successfully.")

}

type SafetyItem struct {
	Name     string `json:"name"`
	Describe string `json:"describe"`
}

func export_docx() {
	readFile, err := os.Open("word/测试表格.docx")
	if err != nil {
		log.Println(err)
	}
	fileinfo, err := readFile.Stat()
	if err != nil {
		log.Println(err)
	}
	size := fileinfo.Size()
	doc, err := docx.Parse(readFile, size)

	if err != nil {
		log.Println(err)
	}

	for _, it := range doc.Document.Body.Items {
		log.Println(it)
	}

}

func main() {
	export_docx()

	// login_payload := map[string]string{
	// 	"email":    "aipaw@qq.com",
	// 	"password": "Ab123456..",
	// }

	// sendtext_payload := map[string]string{
	// 	"dialog_id": "9146",
	// 	"text":      "你好GPT你是什么时候诞生的",
	// }
	// token, _ := Dootask_login(login_payload)
	// token := "YIG8ANC8q2RG-5ub0yHp4MjO_m7sWTou6BBQwptwp1bRHgCx59JgmzSFAlUgvb73GyAl4fq1NuZtx2hJ54Bz5efdnBWSgnT1bwsdWPlCS4GXqXfevXjfOmC9YPRtZsJf"

	// // dialog_id, _ := Dootask_sendtext(token, sendtext_payload)
	// one_payload := map[string]float64{
	// 	"dialog_id": 678174,
	// }
	// time.Sleep(time.Second * 3)
	// content, _ := Dootask_one(token, one_payload)

	// convert_Md("example.md", content)

}

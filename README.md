# scheduler_prac

1. MS WORD 파일의 템플릿 기반으로 script 를 통해 날짜를 변경해 파일 생성
2. 자동 스케쥴러를 통해 지정한 주기(매일) 파일 생성하도록 스크립트 동작
3. 생성된 파일을 지정한 e-mail 주소로 자동 발송하도록 처리
4. exe파일 실행 시, 스케줄러가 작동하여 매일 오전 9시에 날짜와 시각이 입력(수정)된 docx파일이 output.docx라는 이름으로 생성되며 지정된 메일로 전송된다.

- scheduler_prac: 코드 폴더
- scheduler: template.docx 파일이 위치함
- scheduler_exe: 실행 파일이 위치함
- cscheduleroutput.docx: 생성된 파일

- pyinstaller --onefile --distpath C:\scheduler_exe\schedu
ler --add-data "c:/scheduler/template.docx;." scheduler.py
scheduler.py가 위치한 폴더에서 실행, scheduler.py를 실행 파일로 변경한다. C:\scheduler_exe\scheduler 폴더에 생성하며 c:/scheduler/template.docx 파일을 포함한다.

- 실행 모습
  
![스크린샷 2025-03-05 142629](https://github.com/user-attachments/assets/77a98e5a-02e9-4794-8d12-673fa90a420d)
- 결과
  
![스크린샷 2025-03-05 154342](https://github.com/user-attachments/assets/1df1f91e-c0ad-444e-a454-58bddd1ecfed)

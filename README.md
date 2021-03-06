# Renamer

![image](https://user-images.githubusercontent.com/29726529/153334735-0f3e1dca-7c0b-4c36-b8a8-092407ae9a4d.png)

엑셀을 활용해 대량 파일명 변경/이동 등을 수행하는 프로그램. 파이썬과 `pyinstaller`로 제작됨.

1. 파일명 변경을 수행하고자 하는 폴더에 `makelist.exe` 와 `renamer.exe` 를 압축 풀어놓는다.
2. `makelist.exe` 를 실행하면, 현재 폴더와 하위 폴더의 파일/폴더 목록을 엑셀파일(`renamer.xlsx`)로 추출해준다.
3. 엑셀 파일에서 파일/폴더 별로 변경하고 싶은 이름, 경로 등을 입력하고 저장
4. `renamer.exe` 를 실행하면 그대로 바꿔준다.

### 다운로드

우측의 Release 패널을 눌러 zip파일을 다운 받습니다.

### 규칙

- 파일
    1. 이름을 변경하면, 파일명을 바꿔준다.
    2. 경로를 변경하면, 파일을 다른 곳으로 이동시킨다.
        - 해당 경로가 존재하지 않을 경우, 폴더를 생성한다.
    3. 경로와 이름을 모두 비워두면, 파일을 삭제한다.
        - 복구가 불가능하므로 주의.
- 폴더
    1. 경로를 변경하면, 폴더를 이동시킨다.
        - 해당 경로가 존재하지 않을 경우, 폴더를 생성한다.
    2. 경로를 비워두면, 폴더를 삭제한다.
        - 복구 불가
        - 안에 파일이 있어도 모두 삭제한다.

이외의 상황은 모두 무시된다. ~~아마도~~

### 팁

1. `makelist`에 명령줄 인수(commandline argument)로 경로를 입력하면, 해당 경로와 하위 폴더들의 파일 목록을 현재 위치에 생성한다.
2. `renamer`에 경로 또는 xlsx 파일을 명령줄 인수로 주면, 해당 경로의 `renamer.xlsx` 또는 제공된 xlsx 파일을 근거로 작업을 수행한다.
3. `renamer` 의 명령줄 인수로 `-u` 를 입력하면 `renamer.xlsx` 에 기입된 내용을 역방향으로 실행한다. 즉 작업을 되돌리고 싶을 때 사용.
    1. 삭제한 파일/폴더는 복구 안됨.


### 부끄러운 점

1. `pyinstaller`를 사용한 단일 파일이 Windows Defender에 의해 진단되는 경우가 있다. 예외 허용해줘야 합니다.

### Version history

- ver 1.0:
    릴리즈
- ver 1.1:
    `renamer.exe`가 외부의 경로, 또는 `renamer.xlsx` 파일을 읽을 수 있도록 명령줄 인수 지원.
    엑셀을 읽어들일 때, 수식이 아닌 값을 기준으로 읽어들이도록 변경.    

# WPF PPTX Editor

<img width="2850" height="1504" alt="image" src="https://github.com/user-attachments/assets/bf94b141-d3fd-4330-b311-1869d26c7e21" />

WPF 기반 PowerPoint(.pptx) 편집기입니다. 도형 추가/편집, 연결선 그리기, 텍스트 입력을 지원하며 저장한 파일을 PowerPoint에서 바로 열 수 있습니다.

## Features

- **도형 그리기** — 사각형, 둥근 사각형, 타원, 삼각형, 마름모
- **연결선(Connector)** — 선 그리기 및 도형의 8개 핸들 포인트에 스냅 연결
- **텍스트 박스** — 도형 안에 텍스트 입력, 폰트 크기·정렬 설정
- **드래그 이동** — 도형 드래그, 연결된 선 자동 업데이트
- **핸들 리사이즈** — 도형 8방향 크기 조절, 선 양 끝점 개별 조절
- **스타일 설정** — 채우기 색상, 테두리 색상, 선 두께
- **실행 취소(Undo)** — Ctrl+Z
- **PPTX 저장/열기** — PowerPoint 호환 포맷으로 저장 (OpenXML SDK)
- **슬라이드 추가** — 다중 슬라이드 지원

## Tech Stack

- .NET 8 / WPF
- [CommunityToolkit.Mvvm](https://github.com/CommunityToolkit/dotnet) — MVVM 패턴
- [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) — PPTX 파일 생성

## Getting Started

```bash
git clone https://github.com/kyeongrokkim/wpf-pptx-editor.git
cd wpf-pptx-editor
dotnet run --project wpf-pptx-editor/wpf-pptx-editor.csproj
```

## Requirements

- Windows 10/11
- .NET 8 SDK


Get sheet stats v0.1

소개
 - 엑셀 파일에서 할 수 없는 통계용 DB query 작업을 access를 활용하여 출력
 - 출력 데이터 : 중복, 명_공백, 명_x, 출생년도_공백, 간지_공백, 간지_x

상세정보
 * 테이블과 쿼리가 정의된 access파일이 필요
 1. xlsx파일의 데이터를 access로 import
 2. access에 미리 정의된 sql query를 실행, 결과를 배열에 저장
 3. 저장된 배열을 .txt파일로 출력

- ModelとViewModelのロードについて(2020/10/21)
  Modelはロード可能だが、ViewModelはロード不可能だと思われる
  (ViewModelオブジェクト自体を直接上書きしてしまうと、Viewが更新されなくなる)
  このため、Modelは直接ロードし、ViewModel.Initialize()を通じて、メンバ単位で上書きを行う

- ModelとViewModelのSetup()について(2020/10/21)
  Model.Setup()は、プロジェクト新規作成時のみ実行され、Model.Dataにプロジェクトモデルの
  セットを行う。2度目以降はモデルをロードすればよいが、Model.Dataはオブジェクト型のため
  JsonObjectからObject型への変換が必要となるため、Initialize()を通じて変換する
  ViewModel.Setup()はプロジェクト新規作成時、2度目以降のプロジェクト起動時も実行する

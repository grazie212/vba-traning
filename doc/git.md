# git command

1. カレントディレクトリの指定

        git add .

2. 全てのファイルを登録

        git add --all

3. リポジトリの作成

        git init

4. コミット

        git commit -m 'message'

5. ワーキングツリーからインデックスに登録し，コミットする

        git commit -am 'message'

6. コミットメッセージの変更

        git commit --amend

7. 指定のファイルの差分

        git diff filename

1. ブランチやコミット間の差分

        git diff ブランチ1  ブランチ2

1. checkoutしてるブランチとインデックスとの差分

        git diff --cached

1. コミットログを出力する

        git log

1. コミットログを要約して出力する

        git shortlog

1. コミットログを1行で出力する

        git log --oneline

1. repositoryの状態を表示

        git status

1. 移動・変更
        
        git mv

1. 削除

        git rm

1. HEADの参照とインデックス，ワーキングツリーの状態を戻す

        git reset --hard

1. HEADの参照のみを変更する

        git reset --soft

1. HEADの参照とインデックスを変更する.git resetのデフォルト

        git reset -mixed

1. リポジトリで管理されていないファイルの削除

        git clean -f

1. 削除されるファイルの確認

        git clean -n

1. ディレクトリの削除

        git clean -d

1. repositoryをクローン

        git clone

1. 登録

        git remote add リポジトリ名 パス

1. 更新

        git remote update
1. 特定のリポジトリの更新

        git remote update リポジトリ名

1. 削除

        git remote rm リポジトリ名

1. 存在しなくなったブランチの削除

        git remote prune

1. リモートリポジトリとブランチを指定して反映

        git pull origin (master/develop)

1. リポジトリが登録済みの状態でブランチを指定

        git fetch origin

1. ブランチを指定してマージ

        git merge

1. コミットメッセージを指定してマージ

        git merge ブランチ名 -m 'message'

1. リモートリポジトリ「origin」に「master」ブランチを送信

        git push origin ローカルブランチ名

1. 反映先のブランチを指定

        git push origin ローカルブランチ名:リモートブランチ名

1. ローカルリポジトリに作成

        git branch ブランチ名

1. ローカルブランチの確認

        git branch 

1. リモートトラッキングブランチの確認

        git branch -r

1. ブランチ毎の最新コミットの確認

        git branch -v

1. ブランチ名の変更

        git branch -m ブランチ名(old) ブランチ名(new)

1. ブランチの削除

        git branch -d ブランチ名

1. リモートブランチの削除(ローカルブランチを削除した後)

        git push origin :ブランチ名

1. ブランチのチェックアウト

        git checkout ブランチ名

1. タグ一覧

        git tag

1. タグを付ける

        git tag タグ名 リビジョン

1. タグにメッセージを付ける

        git tag -m 'message'

1. リモートに反映

        git push origin タグ名

1. タグの削除

        git tag -d タグ名

1. 直近のタグの表示

        git describe

1. 指定したブランチにチェックアウトしているブランチのコミットを追従させる

        git rebase ブランチ名


#总是在幽幽暗暗中鬼混，才明白平平淡淡是最真；
Remove-Variable * -ErrorAction SilentlyContinue;$error.Clear()



$dllpath='C:\Windows\System32\MySql.Data.dll'

$today=Get-Date -Format "yyyy-MM-dd HH:mm:ss"

$adminmail='humin@china-xinghan.com'

$htmltab=@"

<style>
        table {
          font-family: arial, sans-serif;
          border-collapse: collapse;
          width: 100%;
        }

         th {
          border: 1px solid #aaaaaa;
          text-align: center;
          padding: 8px;
          background-color: #555555;
          color:#ffffff;
          
        }

         td{
          border: 1px solid #aaaaaa;
          text-align: left;
          padding: 8px;
          max-width:200px;
          min-width:80px;
          font-size:12px;
          color:#333333;
        }

        tr:nth-child(even) {
          background-color: #dddddd;
        }

}
</style>

   

"@




#邮件传递函数
    $sendermesg=@{
    
                    From = 'IT <it@china-xinghan.com>'

                    Cc = $adminmail

                    SmtpServer = '10.9.254.105'
    }

     

#开启数据库连接

  
 
    
                        [void][system.Reflection.Assembly]::LoadWithPartialName("MySql.Data")

                        [void][system.Reflection.Assembly]::LoadFrom($dllpath)

                        $Server="10.9.254.202"

                        $Database="IT"

                        $user="it"

                        $Password= "a*999999"

                        $charset="utf8"

                        $connectionString = "server=$Server;uid=$user;pwd=$Password;database=$Database;charset=$charset"

                        $connection = New-Object MySql.Data.MySqlClient.MySqlConnection($connectionString)

                        $error.Clear()


#开始将AD账户信息写入数据库


        #开启数据库连接并且载入active directory模块

        $connection.Open()

        Import-Module ActiveDirectory

        $aduserlist=Get-ADUser -Filter * -SearchBase "dc=china-xinghan,dc=com" | where {$_.SamAccountName -notmatch "krb*|Guest|Default*"} | Select-Object SamAccountName -ExpandProperty SamAccountName

        foreach ($user in $aduserlist) {

                    $attribute=Get-ADUser -Identity $user -Properties SamAccountName,DisplayName,Enabled,lockedout,Department,Manager,Title,mail,mobile,whenCreated,whenChanged,LastLogonDate,PasswordLastSet,PasswordNeverExpires,msDS-UserPasswordExpiryTimeComputed | Select-Object SamAccountName,DisplayName,Enabled,lockedout,Department,Manager,Title,mail,mobile,whenCreated,whenChanged,LastLogonDate,PasswordLastSet,PasswordNeverExpires,@{Name="PasswordExpirydate";  Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}

                    $memberof=Get-ADPrincipalGroupMembership $user | Select-Object name -ExpandProperty name

                    $SamAccountName=$attribute.SamAccountName

                    $DisplayName=$attribute.DisplayName

                    $mail=$attribute.mail

                    $mobile=$attribute.mobile

                    $lockedout=$attribute.lockedout

                    $PasswordNeverExpires=$attribute.PasswordNeverExpires

                    $Enabled=$attribute.Enabled
                    

                    $department=$attribute.Department

                    $manager=$attribute.Manager

                    if($manager){
                    
                        $managermail=Get-ADUser -Identity $manager -Properties name,mail | Select-Object mail -ExpandProperty mail
                    }else{
                        $managermail=""
                    }

                    $whenCreated=$attribute.whenCreated.ToString("yyyy-MM-dd HH:mm:ss")
             
                    $whenChanged=$attribute.whenChanged.ToString("yyyy-MM-dd HH:mm:ss")

                    $Title=$attribute.Title

                    #最后登陆时间和密码最后设置时间可能为空，需要先判断是否为空，否则插入的是上一笔记录的时间

                    if($attribute.LastLogonDate)
                    {
                    $LastLogonDate=$attribute.LastLogonDate.ToString("yyyy-MM-dd")
                    }
                    else{
                    $LastLogonDate="0000-00-00"
                    }



                    if($attribute.PasswordLastSet)
                    {
                     $PasswordLastSet=$attribute.PasswordLastSet.ToString("yyyy-MM-dd")
                    }
                    else{
                    $PasswordLastSet="0000-00-00"
                    }


                    if($attribute.PasswordExpirydate)
                    {
                     $PasswordExpirydate=$attribute.PasswordExpirydate.ToString("yyyy-MM-dd")
                     $pwdexpiryday=New-TimeSpan  $(Get-Date) $PasswordExpirydate | Select-Object days -ExpandProperty days
                    }
                    else{
                    $PasswordExpirydate="9999-09-09"
                    $pwdexpiryday="9999"
                    }
                   
                   
                   

                    #信息开始写入数据库

                    $insertsql = "INSERT INTO account_aduser_pd(
                                    date,
                                    samaccountname,
                                    displayname,
                                    enabled,
                                    lockedout,
                                    department,
                                    title,
                                    mail,
                                    mobile,
                                    manager,
                                    managermail,
                                    whencreated,
                                    whenchanged,
                                    lastlogondate,
                                    passwordlastset,
                                    PasswordNeverExpires,
                                    passwordexpirydate,
                                    pwdexpiryday,
                                    memberof
                                    )
                                    VALUES(
                                    '$today',
                                    '$SamAccountName',
                                    '$DisplayName',
                                    '$Enabled',
                                    '$lockedout',
                                    '$Department',
                                    '$Title',
                                    '$mail',
                                    '$mobile',
                                    '$Manager',
                                    '$managermail',
                                    '$whenCreated',
                                    '$whenChanged',
                                    '$LastLogonDate',
                                    '$PasswordLastSet',
                                    '$PasswordNeverExpires',
                                    '$PasswordExpirydate',
                                    '$pwdexpiryday',
                                    '$memberof'
                                    )"

                    $insertcommand = New-Object MySql.Data.MySqlClient.MySqlCommand

                    $insertcommand.Connection=$connection

                    $insertcommand.CommandText=$insertsql

                    $insertcommand.ExecuteNonQuery()

}

    $connection.Close()
        

#开始查询系统加固并生成生产报告，邮件通知

            try{

                    $connection.Open()

                    $scSelectComm = New-Object MySql.Data.MySqlClient.MySqlCommand

                    $scSelectComm.Connection=$connection

                    $scSelectComm.CommandText='SELECT area,hostname,ip,operasystem,date,result FROM systemcheck_report 
                                                    WHERE
                                                     area ="IT基础架构" OR 
                                                     area ="DataBuffer" OR
                                                     area ="生产服务器" OR
                                                     area ="生产客户端" OR
                                                     area ="监控室" AND
                                                     TO_DAYS(NOW())-TO_DAYS(date)=1 GROUP BY hostname'

                    $adaptersc = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($scSelectComm)

                    $datasetsc = New-Object System.Data.DataSet

                    $adaptersc.Fill($datasetsc)


                    $rowcount = $datasetsc.Tables[0].Rows.Count

                    $rowline=0

 
                                foreach($row in $datasetsc.Tables[0].Rows){
                               

                                    $rowline ++

                                    $area = $row["area"]

                                    $hostname = $row["hostname"]
                                   
                                    $ip = $row["ip"]
                                   
                                    $operasystem = $row["operasystem"]

                                    $result = $row["result"]

                                    $date = $row["date"]

                                    $checklist +="<tr><td>$rowline</td><td>$area</td><td> $hostname </td><td> $ip </td><td> $operasystem </td><td> $result </td><td> $date </td></tr>"

                                    
                                }
                               
                                $sendermesg.Body =$htmltab+"<h3> 今日已报告主机清单( $rowcount 台)</h3><table><tr><th>序号</th><th>网络区域</th><th>主机名</th><th>ip地址</th><th>操作系统</th><th>加固状态</th><th>日期时间</th></tr>"+$checklist+"</table>"

                                $sendermesg.Subject="【日志邮件通知】：生产网系统加固检查报告 $date"

                                $sendermesg.To = "humin@china-xinghan.com"

                                Send-MailMessage @sendermesg -Encoding 'utf8' -BodyAsHtml

                                
                     $connection.Close()

                     
                
                        
            }
            catch{

            Write-Output "系统加固Error: $($PSItem.ToString())" | out-file -filepath C:\ProgramData\xinghan.txt

            }



#开始查询AD账户信息：密码过期状态、为登录信息状态并通知

            try{

               

                $connection.Open()

                $selectcommand = New-Object MySql.Data.MySqlClient.MySqlCommand

                $selectcommand.Connection=$connection

                $selectcommand.CommandText="SELECT * FROM account_aduser_pd WHERE DATE(date)=CURDATE() and mobile not like '服务账户' GROUP BY samaccountname"
                 
                $adapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($selectcommand)

                $dataset = New-Object System.Data.DataSet

                $adapter.Fill($dataset)

                        $rowline=0

                        foreach ($row in $dataset.Tables[0].Rows)
                        
                         {
                            $rowline ++

                            $name = $row["samaccountname"]

                            $accounttype = $row["accounttype"]

                            $displayname = $row["displayname"]

                            $enabled = $row["enabled"]

                            $department = $row["department"]

                            $title = $row["title"]

                            $mail = $row["mail"]

                            $lockedout = $row["lockedout"]

                            $manager = $row["manager"]

                            $managermail = $row["managermail"]

                            $whencreated = $row["whencreated"]

                            $lastlogondate = ($row["lastlogondate"]).ToString("yyyy-MM-dd")

                            $passwordlastset = ($row["passwordlastset"]).ToString("yyyy-MM-dd")

                            $passwordexpirydate = ($row["passwordexpirydate"]).ToString("yyyy-MM-dd")

                            $nologondate=(New-TimeSpan $row["lastlogondate"] -End (Get-Date -Format ("yyyy-MM-dd"))).Days  

                            $date = $row["date"]

                            $pwdexpiryday = $row["pwdexpiryday"]


                                      if($pwdexpiryday -le 5)
                                                {

                                                    try{
                                                    
                                                                switch($pwdexpiryday)

                                                                                  {
                              
                                                                                    {$PSItem -le 5 -and $PSItem -ge 1 }{ $sendermesg.Subject="$displayname 域控账户密码即将在 $pwdexpiryday 天后过期，请尽快修改！" }
                                                                                    {$PSItem -lt 1}{ $sendermesg.Subject="$displayname 域控账户密码已过期 "+-$pwdexpiryday+" 天，请立即修改！"}
                                                              
                                                                                  }
                                    
                                                                     $sendermesg.Body=$htmltab+"<table><tr><th>账户</th><th>姓名</th><th>上一次密码修改日期</th><th>密码到期日期</th><th>密码过期天数</th></tr><tr><td>$name</td><td> $displayname </td><td> $passwordlastset </td><td> $passwordexpirydate </td><td> $pwdexpiryday </td></tr></table>"
                                                         
                                                                     if($mail){
                                                 
                                                                         $sendermesg.To = "humin@china-xinghan.com"

                                                                         }
                                                                         else{
                                                 
                                                     
                                                                         $sendermesg.To = "humin@china-xinghan.com"

                                                                     }
                                                 
                                                                     Send-MailMessage @sendermesg -Encoding 'utf8' -BodyAsHtml

                                                                     sleep(15)
                                                    }
                                                    catch{

                                                                    Write-Output "Error: $($PSItem.ToString())" | out-file -filepath C:\ProgramData\xinghan.txt
                                                    
                                                    }
                                    
                                                        

                                    
                                         }
                                         
                                         if($nologondate -gt 15)

                                                    {

                                                    try{
                                                                #查询超过15天没有登录的用户
                                                                        $sendermesg.Subject="$displayname ,你的域控账户在 $nologondate 天内未登录过，不符合安全认证要求！"
                                                    
                                                                        $sendermesg.Body=$htmltab+"<table><tr><th>账户</th><th>姓名</th><th>上一次登录时间</th><th>未登录时长（天）</th></tr><tr><td>$name</td><td> $displayname </td><td> $lastlogondate </td><td> $nologondate </td></tr></table>"
                                                           
                                                                        if($mail){
                                                 
                                                                             $sendermesg.To = "humin@china-xinghan.com"

                                                                             }
                                                                             else{
                                                 
                                                                             $sendermesg.To = "humin@china-xinghan.com"

                                                                        }
                                                                        Send-MailMessage @sendermesg -Encoding 'utf8' -BodyAsHtml
                                                    
                                                    }

                                                    catch
                                                    {
                                                                
                                                                Write-Output "Error: $($PSItem.ToString())" | out-file -filepath C:\ProgramData\xinghan.txt
                                                    
                                                    }
                                    
                                                            
                                    
                                    
                                                }


                                    

                             
                             
                             
                              $userlist +="<tr>
                                    <td> $rowline</td>
                                    <td> $name</td>
                                    <td> $displayname </td>
                                    <td> $accounttype </td>
                                    <td> $enabled </td>
                                    <td> $department </td>
                                    <td> $whencreated </td>
                                    <td> $lastlogondate </td>
                                    <td> $passwordlastset</td>
                                    <td> $passwordexpirydate </td>
                                    <td> $pwdexpiryday </td>
                                    <td> $date </td></tr>"     
                                                    
                                   
                          }
                              
                           $sendermesg.Subject="AD账户明细清单（生产网） $today" 
                                   
                           $sendermesg.Body =$htmltab+"<h3> AD账户明细清单（生产网）</h3><table><tr><th>序号</th><th>账户ID</th><th>账户名</th><th>账户类型</th><th>账户状态</th><th>部门</th><th>创建时间</th><th>最近登录日期</th><th>密码修改日期</th><th>密码到期日期</th><th>密码到期天数</th><th>报告日期</th></tr>"+$userlist+"</table>"
                               
                           $sendermesg.To = "humin@china-xinghan.com"
                           
                           Send-MailMessage @sendermesg -Encoding 'utf8' -BodyAsHtml


                     $connection.Close()     

                        
                        

            }

            catch{

                    Write-Output "Error: $($PSItem.ToString())" | out-file -filepath C:\ProgramData\xinghan.txt
            }
            
 
pipeline {
    agent any

    stages {
        stage('Checkout') {
            steps {
                checkout scm
            }
        }

        stage('Clone Repository') {
            steps {
                git branch: 'main', url: 'https://github.com/Shantanu26-acg/L4-Azure.git'
            }
        }

        stage('Run Tests') {
            steps {
                // sh 'pytest --maxfail=1 --disable-warnings -q'
                // withEnv(["PYTHONENCODING=utf-8"]) {
                //     // bat "\"C:/Users/Shantanu/AppData/Local/Programs/Python/Python313/python.exe\" -m pip install -r requirements.txt"
                //     // bat "\"C:/Users/Shantanu/AppData/Local/Programs/Python/Python313/python.exe\" Create_role_new.py -v --html=report.html"
                //     // bat "\"C:/Users/Shantanu/AppData/Local/Programs/Python/Python313/python.exe\" Create_User_new.py -v --html=report.html"
                //     // bat "\"C:/Users/Shantanu/AppData/Local/Programs/Python/Python313/python.exe\" Create_Location_new.py -v --html=report.html"
                //     // bat "\"C:/Users/Shantanu/AppData/Local/Programs/Python/Python313/python.exe\" Create_Product_SNG.py -v --html=report.html"
                //     // bat "\"C:/Users/Shantanu/AppData/Local/Programs/Python/Python313/python.exe\" Create_SSCC_Template_new.py -v --html=report.html"
                // }
                script{
                    def scripts = [
                        "Create_role_new.py",
                        "Create_User_new.py",
                        "Create_Location_new.py",
                        "Create_Product_SNG.py",
                        "Create_SSCC_Template_new.py"
                    ]
                    
                    for (s in scripts) {
                        catchError(buildResult: 'SUCCESS', stageResult: 'FAILURE') {
                            bat "\"C:/Users/Shantanu/AppData/Local/Programs/Python/Python313/python.exe\" ${s} -v --html=report.html"
                        }
                    }
                }
            }
        }
    }

    post {
        always {
            // junit '**/pytest.xml'
                archiveArtifacts artifacts: 'report.html', fingerprint: true
        }
    }
}

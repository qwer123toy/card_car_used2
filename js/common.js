// 공통 Javascript 함수 모음

// 문서 로드 완료 후 실행
document.addEventListener('DOMContentLoaded', function() {
    // 알림 메시지가 있는 경우 자동으로 사라지도록 설정
    const alertElements = document.querySelectorAll('.alert');
    if (alertElements.length > 0) {
        alertElements.forEach(function(alert) {
            setTimeout(function() {
                alert.style.opacity = '0';
                alert.style.transition = 'opacity 0.5s';
                setTimeout(function() {
                    alert.remove();
                }, 500);
            }, 3000);
        });
    }
});

// 폼 유효성 검사
function validateForm(formId, rules) {
    const form = document.getElementById(formId);
    
    if (!form) return true;
    
    let isValid = true;
    
    for (const field in rules) {
        const input = form.querySelector(`[name="${field}"]`);
        const value = input.value.trim();
        const rule = rules[field];
        
        // 필수 입력 검사
        if (rule.required && value === '') {
            showError(input, rule.message || '필수 입력 항목입니다.');
            isValid = false;
            continue;
        }
        
        // 최소 길이 검사
        if (rule.minLength && value.length < rule.minLength) {
            showError(input, rule.message || `최소 ${rule.minLength}자 이상 입력해야 합니다.`);
            isValid = false;
            continue;
        }
        
        // 이메일 형식 검사
        if (rule.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value)) {
            showError(input, rule.message || '올바른 이메일 형식이 아닙니다.');
            isValid = false;
            continue;
        }
        
        // 숫자만 검사
        if (rule.numeric && !/^\d+$/.test(value)) {
            showError(input, rule.message || '숫자만 입력할 수 있습니다.');
            isValid = false;
            continue;
        }
        
        // 패턴 검사
        if (rule.pattern && !rule.pattern.test(value)) {
            showError(input, rule.message || '형식이 올바르지 않습니다.');
            isValid = false;
            continue;
        }
        
        // 에러 메시지 제거
        clearError(input);
    }
    
    return isValid;
}

// 입력창 에러 표시
function showError(input, message) {
    const formGroup = input.closest('.form-group');
    const errorElement = formGroup.querySelector('.error-message') || document.createElement('div');
    
    if (!formGroup.querySelector('.error-message')) {
        errorElement.className = 'error-message';
        errorElement.style.color = 'red';
        errorElement.style.fontSize = '0.8rem';
        errorElement.style.marginTop = '0.25rem';
        formGroup.appendChild(errorElement);
    }
    
    errorElement.textContent = message;
    input.style.borderColor = 'red';
}

// 에러 메시지 제거
function clearError(input) {
    const formGroup = input.closest('.form-group');
    const errorElement = formGroup.querySelector('.error-message');
    
    if (errorElement) {
        errorElement.remove();
    }
    
    input.style.borderColor = '';
}

// 금액 형식 설정 (천단위 콤마)
function formatCurrency(input) {
    let value = input.value.replace(/,/g, '');
    
    if (value) {
        value = parseInt(value, 10).toLocaleString('ko-KR');
        input.value = value;
    }
}

// 날짜 형식 설정 (YYYY-MM-DD)
function formatDate(input) {
    const value = input.value.replace(/\D/g, '');
    if (value.length >= 8) {
        const year = value.substring(0, 4);
        const month = value.substring(4, 6);
        const day = value.substring(6, 8);
        input.value = `${year}-${month}-${day}`;
    }
}

// AJAX 요청 함수
function ajaxRequest(options) {
    const xhr = new XMLHttpRequest();
    
    xhr.open(options.method || 'GET', options.url, true);
    
    if (options.headers) {
        Object.keys(options.headers).forEach(key => {
            xhr.setRequestHeader(key, options.headers[key]);
        });
    }
    
    xhr.onload = function() {
        if (xhr.status >= 200 && xhr.status < 300) {
            if (options.success) {
                try {
                    const response = JSON.parse(xhr.responseText);
                    options.success(response);
                } catch (e) {
                    options.success(xhr.responseText);
                }
            }
        } else {
            if (options.error) {
                options.error(xhr.statusText, xhr.status);
            }
        }
    };
    
    xhr.onerror = function() {
        if (options.error) {
            options.error('네트워크 오류가 발생했습니다.');
        }
    };
    
    xhr.send(options.data || null);
} 
import { async, ComponentFixture, TestBed } from '@angular/core/testing';

import { EditdilalogcmpComponent } from './editdilalogcmp.component';

describe('EditdilalogcmpComponent', () => {
  let component: EditdilalogcmpComponent;
  let fixture: ComponentFixture<EditdilalogcmpComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [ EditdilalogcmpComponent ]
    })
    .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(EditdilalogcmpComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});

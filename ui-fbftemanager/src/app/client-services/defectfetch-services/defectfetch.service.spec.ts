import { TestBed } from '@angular/core/testing';

import { DefectfetchService } from './defectfetch.service';

describe('DefectfetchService', () => {
  beforeEach(() => TestBed.configureTestingModule({}));

  it('should be created', () => {
    const service: DefectfetchService = TestBed.get(DefectfetchService);
    expect(service).toBeTruthy();
  });
});
